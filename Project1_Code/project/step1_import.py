"""
Step 1 – Data Preparation
Import Reviews_withURL.csv into MongoDB collection 'reviews'
inside database 'amazon_reviews'.
"""

import pandas as pd
from pymongo import MongoClient, UpdateOne
import time

CSV_PATH = "./Reviews_withURL.csv"
MONGO_URI = "mongodb://localhost:27017"
DB_NAME   = "amazon_reviews"
COL_NAME  = "reviews"

print("Loading CSV …")
df = pd.read_csv(CSV_PATH, index_col=0)          # first col is unnamed row-index
print(f"  Rows: {len(df):,}  |  Columns: {list(df.columns)}")

# ── Clean up ──────────────────────────────────────────────────────────────────
# Convert timestamp (Unix seconds) to datetime
df["ReviewDate"] = pd.to_datetime(df["Time"], unit="s")

# Derive a new field: SentimentLabel  (Score 4-5 → Positive, 3 → Neutral, 1-2 → Negative)
def label(score):
    if score >= 4:  return "Positive"
    if score == 3:  return "Neutral"
    return "Negative"

df["SentimentLabel"] = df["Score"].apply(label)

# Helpfulness ratio (avoid division by zero)
df["HelpfulnessRatio"] = df.apply(
    lambda r: round(r["HelpfulnessNumerator"] / r["HelpfulnessDenominator"], 4)
    if r["HelpfulnessDenominator"] > 0 else None,
    axis=1
)

records = df.to_dict(orient="records")

# ── Import ────────────────────────────────────────────────────────────────────
client = MongoClient(MONGO_URI)
db     = client[DB_NAME]
col    = db[COL_NAME]

col.drop()          # fresh import each run
print("Inserting documents …")
t0 = time.time()
BATCH = 10_000
for i in range(0, len(records), BATCH):
    col.insert_many(records[i:i+BATCH])
    if (i // BATCH) % 10 == 0:
        print(f"  {i+BATCH:,} / {len(records):,} inserted …")

elapsed = time.time() - t0
print(f"\nImport complete in {elapsed:.1f}s")
print(f"  Total documents: {col.count_documents({}):,}")

# ── Indexes ───────────────────────────────────────────────────────────────────
col.create_index("ProductId")
col.create_index("UserId")
col.create_index("Score")
col.create_index("SentimentLabel")
col.create_index("ReviewDate")
print("Indexes created.")
client.close()
