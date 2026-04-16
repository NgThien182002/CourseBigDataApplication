"""
Step 2 – Natural Language Processing
1.  Efficient word-frequency counting with Counter (no NLTK tokeniser overhead)
2.  TF-IDF keyword extraction per sentiment class
3.  Named-entity style keyword extraction using noun-phrase chunking (NLTK)
4.  Bigram analysis
"""

import re, string
from collections import Counter
import pandas as pd
from pymongo import MongoClient
import nltk

# Download once
for pkg in ["punkt", "stopwords", "averaged_perceptron_tagger", "punkt_tab",
            "averaged_perceptron_tagger_eng"]:
    nltk.download(pkg, quiet=True)

from nltk.corpus import stopwords
from nltk import word_tokenize, pos_tag, ne_chunk, ngrams
from sklearn.feature_extraction.text import TfidfVectorizer
import json, os

MONGO_URI = "mongodb://localhost:27017"
DB_NAME   = "amazon_reviews"
COL_NAME  = "reviews"
OUT_DIR   = "./nlp_results"
os.makedirs(OUT_DIR, exist_ok=True)

# ─── helpers ──────────────────────────────────────────────────────────────────
STOP = set(stopwords.words("english"))
STOP |= {"br", "amazon", "product", "one", "would", "get", "also", "like", "much",
          "really", "even", "made", "make", "use", "used", "great", "good", "food",
          "item", "buy", "bought", "still", "taste", "flavor", "well", "got", "back"}

def clean_tokens(text: str) -> list:
    text = text.lower()
    text = re.sub(r"<.*?>", " ", text)          # strip HTML tags
    text = re.sub(r"[^a-z\s]", " ", text)       # keep only letters
    tokens = text.split()
    return [t for t in tokens if t not in STOP and len(t) > 2]

# ─── load data from MongoDB ───────────────────────────────────────────────────
print("Connecting to MongoDB …")
client = MongoClient(MONGO_URI)
col    = client[DB_NAME][COL_NAME]

print("Fetching reviews …")
cursor = col.find({}, {"Text": 1, "Summary": 1, "Score": 1, "SentimentLabel": 1, "_id": 0})
rows   = list(cursor)
client.close()
print(f"  Loaded {len(rows):,} documents")

df = pd.DataFrame(rows)
df["combined"] = df["Summary"].fillna("") + " " + df["Text"].fillna("")

# ════════════════════════════════════════════════════════════════════════════════
# TASK A – Overall word frequency (most efficient: single Counter pass)
# ════════════════════════════════════════════════════════════════════════════════
print("\n[A] Counting word frequencies …")
total_counter: Counter = Counter()
for text in df["combined"]:
    total_counter.update(clean_tokens(text))

top100 = total_counter.most_common(100)
wf_df  = pd.DataFrame(top100, columns=["word", "count"])
wf_df.to_csv(f"{OUT_DIR}/word_freq_top100.csv", index=False)
print(f"  Top 10: {top100[:10]}")

# ════════════════════════════════════════════════════════════════════════════════
# TASK B – Word frequency per sentiment (efficient: grouped Counter)
# ════════════════════════════════════════════════════════════════════════════════
print("\n[B] Word freq by sentiment …")
sent_counters = {}
for label, group in df.groupby("SentimentLabel"):
    c = Counter()
    for text in group["combined"]:
        c.update(clean_tokens(text))
    sent_counters[label] = c
    top5 = c.most_common(5)
    print(f"  {label}: {top5}")
    pd.DataFrame(c.most_common(50), columns=["word","count"]).to_csv(
        f"{OUT_DIR}/word_freq_{label.lower()}.csv", index=False)

# ════════════════════════════════════════════════════════════════════════════════
# TASK C – TF-IDF keyword extraction per sentiment class
# ════════════════════════════════════════════════════════════════════════════════
print("\n[C] TF-IDF per sentiment …")
sentiment_docs = {}
for label, group in df.groupby("SentimentLabel"):
    # Concatenate all reviews in that class into one large document
    sentiment_docs[label] = " ".join(group["combined"].fillna("").tolist())

tfidf = TfidfVectorizer(max_features=200, stop_words="english",
                        ngram_range=(1, 2), min_df=1)
tfidf_matrix = tfidf.fit_transform(list(sentiment_docs.values()))
feature_names = tfidf.get_feature_names_out()
labels_list   = list(sentiment_docs.keys())

tfidf_results = {}
for i, label in enumerate(labels_list):
    scores = tfidf_matrix[i].toarray().flatten()
    top_idx = scores.argsort()[::-1][:20]
    top_keywords = [(feature_names[j], round(float(scores[j]), 4)) for j in top_idx]
    tfidf_results[label] = top_keywords
    print(f"  {label} top keywords: {top_keywords[:5]}")

with open(f"{OUT_DIR}/tfidf_keywords.json", "w") as f:
    json.dump(tfidf_results, f, indent=2)

# ════════════════════════════════════════════════════════════════════════════════
# TASK D – Bigram analysis (top 50 bigrams overall)
# ════════════════════════════════════════════════════════════════════════════════
print("\n[D] Bigram analysis …")
bigram_counter: Counter = Counter()
SAMPLE_N = 50_000   # use a sample for speed
sample = df["combined"].sample(n=SAMPLE_N, random_state=42)
for text in sample:
    tokens = clean_tokens(text)
    bigram_counter.update(ngrams(tokens, 2))

top_bigrams = [(f"{a} {b}", c) for (a, b), c in bigram_counter.most_common(50)]
pd.DataFrame(top_bigrams, columns=["bigram", "count"]).to_csv(
    f"{OUT_DIR}/bigrams_top50.csv", index=False)
print(f"  Top 10 bigrams: {top_bigrams[:10]}")

# ════════════════════════════════════════════════════════════════════════════════
# TASK E – Demonstrate searching for a specific word in MongoDB (query approach)
# ════════════════════════════════════════════════════════════════════════════════
SEARCH_WORD = "delicious"
print(f"\n[E] Counting appearance of '{SEARCH_WORD}' via regex query …")
client2 = MongoClient(MONGO_URI)
col2    = client2[DB_NAME][COL_NAME]

regex_count = col2.count_documents(
    {"Text": {"$regex": SEARCH_WORD, "$options": "i"}}
)
print(f"  '{SEARCH_WORD}' appears in {regex_count:,} reviews (MongoDB regex query)")

# Also demonstrate text-index approach (most efficient for full-text search)
try:
    col2.create_index([("Text", "text"), ("Summary", "text")])
    text_result = col2.find(
        {"$text": {"$search": SEARCH_WORD}},
        {"score": {"$meta": "textScore"}, "Summary": 1, "Score": 1, "_id": 0}
    ).sort([("score", {"$meta": "textScore"})]).limit(3)
    print("  Text-index top-3 results:")
    for doc in text_result:
        print(f"    Score={doc['Score']} | {doc['Summary'][:60]}")
except Exception as e:
    print(f"  Text index note: {e}")

client2.close()

print("\nStep 2 NLP complete. Results saved to", OUT_DIR)
