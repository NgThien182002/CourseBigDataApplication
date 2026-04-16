"""
Step 3 – Analysis & Visualization
Produces 8 publication-quality charts saved to /project/figures/
"""

import os, json, warnings
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
from pymongo import MongoClient
from collections import Counter
from wordcloud import WordCloud
import re

warnings.filterwarnings("ignore")
FIG_DIR = "./figures"
os.makedirs(FIG_DIR, exist_ok=True)

# ── Palette ───────────────────────────────────────────────────────────────────
PALETTE = {"Positive": "#2ecc71", "Neutral": "#f39c12", "Negative": "#e74c3c"}
BLUE    = "#3498db"
DARK    = "#2c3e50"
sns.set_theme(style="whitegrid", font_scale=1.15)

# ── Load data ─────────────────────────────────────────────────────────────────
print("Loading data from MongoDB …")
client = MongoClient("mongodb://localhost:27017")
col    = client["amazon_reviews"]["reviews"]
cursor = col.find({}, {"Score": 1, "SentimentLabel": 1, "ReviewDate": 1,
                        "HelpfulnessNumerator": 1, "HelpfulnessDenominator": 1,
                        "HelpfulnessRatio": 1, "combined": 1, "Text": 1,
                        "Summary": 1, "_id": 0})
rows = list(cursor)
client.close()

df = pd.DataFrame(rows)
df["ReviewDate"] = pd.to_datetime(df["ReviewDate"], errors="coerce")
df["Year"]  = df["ReviewDate"].dt.year
df["Month"] = df["ReviewDate"].dt.to_period("M").dt.to_timestamp()
df["combined"] = df["Summary"].fillna("") + " " + df["Text"].fillna("")

print(f"  {len(df):,} rows loaded")

# ══════════════════════════════════════════════════════════════════════════════
# FIG 1 – Score Distribution
# ══════════════════════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(8, 5))
score_counts = df["Score"].value_counts().sort_index()
bars = ax.bar(score_counts.index, score_counts.values,
              color=[PALETTE["Negative"], PALETTE["Negative"],
                     PALETTE["Neutral"],  PALETTE["Positive"], PALETTE["Positive"]],
              edgecolor="white", width=0.6)
for bar, val in zip(bars, score_counts.values):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 2000,
            f"{val:,}", ha="center", va="bottom", fontsize=10, color=DARK)
ax.set_xlabel("Star Rating", fontsize=12)
ax.set_ylabel("Number of Reviews", fontsize=12)
ax.set_title("Distribution of Star Ratings", fontsize=14, fontweight="bold", color=DARK)
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{int(x):,}"))
ax.set_xticks([1, 2, 3, 4, 5])
plt.tight_layout()
plt.savefig(f"{FIG_DIR}/fig1_score_distribution.png", dpi=150)
plt.close()
print("  Fig 1 saved")

# ══════════════════════════════════════════════════════════════════════════════
# FIG 2 – Sentiment Pie Chart
# ══════════════════════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(7, 5))
sent_counts = df["SentimentLabel"].value_counts()
colors = [PALETTE[l] for l in sent_counts.index]
wedges, texts, autotexts = ax.pie(
    sent_counts.values, labels=sent_counts.index,
    colors=colors, autopct="%1.1f%%",
    startangle=140, pctdistance=0.75,
    wedgeprops={"edgecolor": "white", "linewidth": 2})
for at in autotexts:
    at.set_fontsize(11); at.set_fontweight("bold")
ax.set_title("Sentiment Distribution", fontsize=14, fontweight="bold", color=DARK)
plt.tight_layout()
plt.savefig(f"{FIG_DIR}/fig2_sentiment_pie.png", dpi=150)
plt.close()
print("  Fig 2 saved")

# ══════════════════════════════════════════════════════════════════════════════
# FIG 3 – Review Volume Over Time (yearly)
# ══════════════════════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(10, 5))
yearly = df.groupby("Year").size().reset_index(name="count")
yearly = yearly[yearly["Year"].between(2000, 2012)]
ax.fill_between(yearly["Year"], yearly["count"], alpha=0.25, color=BLUE)
ax.plot(yearly["Year"], yearly["count"], marker="o", color=BLUE, linewidth=2.5)
for _, row in yearly.iterrows():
    ax.text(row["Year"], row["count"] + 1500, f"{int(row['count']):,}",
            ha="center", fontsize=8.5, color=DARK)
ax.set_xlabel("Year", fontsize=12)
ax.set_ylabel("Number of Reviews", fontsize=12)
ax.set_title("Review Volume Growth (2000 – 2012)", fontsize=14, fontweight="bold", color=DARK)
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{int(x):,}"))
ax.set_xticks(range(2000, 2013))
plt.tight_layout()
plt.savefig(f"{FIG_DIR}/fig3_volume_over_time.png", dpi=150)
plt.close()
print("  Fig 3 saved")

# ══════════════════════════════════════════════════════════════════════════════
# FIG 4 – Average Score Over Time
# ══════════════════════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(10, 5))
avg_score = df[df["Year"].between(2000, 2012)].groupby("Year")["Score"].mean()
ax.plot(avg_score.index, avg_score.values, marker="s", color="#9b59b6",
        linewidth=2.5, markersize=7)
ax.axhline(df["Score"].mean(), linestyle="--", color="grey", alpha=0.6,
           label=f"Overall mean: {df['Score'].mean():.2f}")
ax.set_ylim(1, 5.2)
ax.set_xlabel("Year", fontsize=12)
ax.set_ylabel("Average Star Rating", fontsize=12)
ax.set_title("Average Rating per Year", fontsize=14, fontweight="bold", color=DARK)
ax.set_xticks(range(2000, 2013))
ax.legend()
plt.tight_layout()
plt.savefig(f"{FIG_DIR}/fig4_avg_score_over_time.png", dpi=150)
plt.close()
print("  Fig 4 saved")

# ══════════════════════════════════════════════════════════════════════════════
# FIG 5 – Top 20 Words (bar chart)
# ══════════════════════════════════════════════════════════════════════════════
wf_path = "./nlp_results/word_freq_top100.csv"
if os.path.exists(wf_path):
    wf_df = pd.read_csv(wf_path).head(20)
else:
    # quick fallback
    from collections import Counter
    stop = {"the","and","is","in","it","of","a","to","i","this","have","for",
            "with","was","my","but","are","be","that","not","they","we","so",
            "br","amazon","product","food","good","great"}
    c = Counter()
    for text in df["combined"].sample(50000, random_state=1):
        c.update([w for w in re.sub(r"[^a-z ]", " ", text.lower()).split()
                  if w not in stop and len(w) > 2])
    wf_df = pd.DataFrame(c.most_common(20), columns=["word","count"])

fig, ax = plt.subplots(figsize=(10, 6))
sns.barplot(data=wf_df, y="word", x="count", palette="Blues_r", ax=ax)
ax.set_xlabel("Frequency", fontsize=12)
ax.set_ylabel("Word", fontsize=12)
ax.set_title("Top 20 Most Frequent Words in Reviews", fontsize=14,
             fontweight="bold", color=DARK)
ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{int(x):,}"))
plt.tight_layout()
plt.savefig(f"{FIG_DIR}/fig5_top20_words.png", dpi=150)
plt.close()
print("  Fig 5 saved")

# ══════════════════════════════════════════════════════════════════════════════
# FIG 6 – Word Cloud (positive reviews)
# ══════════════════════════════════════════════════════════════════════════════
pos_path = "./nlp_results/word_freq_positive.csv"
if os.path.exists(pos_path):
    pos_df = pd.read_csv(pos_path)
    freq_dict = dict(zip(pos_df["word"], pos_df["count"]))
else:
    freq_dict = {}

if not freq_dict:
    from collections import Counter
    stop = {"the","and","is","in","it","of","a","to","i","this","have","for",
            "with","was","my","but","are","be","that","not","they","good","great"}
    c = Counter()
    for text in df[df["SentimentLabel"]=="Positive"]["combined"].sample(30000, random_state=1):
        c.update([w for w in re.sub(r"[^a-z ]", " ", text.lower()).split()
                  if w not in stop and len(w) > 2])
    freq_dict = dict(c.most_common(200))

wc = WordCloud(width=1000, height=500, background_color="white",
               colormap="Greens", max_words=150,
               prefer_horizontal=0.9).generate_from_frequencies(freq_dict)
fig, ax = plt.subplots(figsize=(12, 6))
ax.imshow(wc, interpolation="bilinear")
ax.axis("off")
ax.set_title("Word Cloud – Positive Reviews", fontsize=15,
             fontweight="bold", color=DARK, pad=15)
plt.tight_layout()
plt.savefig(f"{FIG_DIR}/fig6_wordcloud_positive.png", dpi=150)
plt.close()
print("  Fig 6 saved")

# ══════════════════════════════════════════════════════════════════════════════
# FIG 7 – Helpfulness Ratio Distribution by Sentiment
# ══════════════════════════════════════════════════════════════════════════════
hr_df = df[df["HelpfulnessRatio"].notna() & (df["HelpfulnessDenominator"] >= 5)]
fig, ax = plt.subplots(figsize=(9, 5))
for lbl, clr in PALETTE.items():
    sub = hr_df[hr_df["SentimentLabel"] == lbl]["HelpfulnessRatio"]
    ax.hist(sub, bins=30, alpha=0.55, color=clr, label=lbl, density=True)
ax.set_xlabel("Helpfulness Ratio (helpful / total votes)", fontsize=12)
ax.set_ylabel("Density", fontsize=12)
ax.set_title("Helpfulness Ratio by Sentiment", fontsize=14,
             fontweight="bold", color=DARK)
ax.legend()
plt.tight_layout()
plt.savefig(f"{FIG_DIR}/fig7_helpfulness_by_sentiment.png", dpi=150)
plt.close()
print("  Fig 7 saved")

# ══════════════════════════════════════════════════════════════════════════════
# FIG 8 – Heatmap: Avg Score by Year × Sentiment Share
# ══════════════════════════════════════════════════════════════════════════════
pivot = (df[df["Year"].between(2002, 2012)]
         .groupby(["Year", "SentimentLabel"])
         .size()
         .unstack(fill_value=0))
pivot_pct = pivot.div(pivot.sum(axis=1), axis=0) * 100

fig, ax = plt.subplots(figsize=(10, 5))
sns.heatmap(pivot_pct[["Positive","Neutral","Negative"]].T,
            annot=True, fmt=".1f", cmap="RdYlGn",
            linewidths=0.5, ax=ax, cbar_kws={"label": "% of Reviews"})
ax.set_xlabel("Year", fontsize=12)
ax.set_ylabel("Sentiment", fontsize=12)
ax.set_title("Sentiment Share by Year (%)", fontsize=14,
             fontweight="bold", color=DARK)
plt.tight_layout()
plt.savefig(f"{FIG_DIR}/fig8_sentiment_heatmap.png", dpi=150)
plt.close()
print("  Fig 8 saved")

print(f"\nAll figures saved to {FIG_DIR}")

# ── Summary Stats (for slide deck) ───────────────────────────────────────────
summary = {
    "total_reviews"    : int(len(df)),
    "total_users"      : int(df["UserId"].nunique()) if "UserId" in df.columns else "N/A",
    "avg_score"        : round(float(df["Score"].mean()), 2),
    "pct_positive"     : round(float((df["SentimentLabel"]=="Positive").mean()*100), 1),
    "pct_neutral"      : round(float((df["SentimentLabel"]=="Neutral").mean()*100), 1),
    "pct_negative"     : round(float((df["SentimentLabel"]=="Negative").mean()*100), 1),
    "date_range"       : "Oct 1999 – Oct 2012",
}
with open("./summary_stats.json", "w") as f:
    json.dump(summary, f, indent=2)
print("Summary stats:", summary)
