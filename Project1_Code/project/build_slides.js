const PptxGenJS = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const pptx = new PptxGenJS();

// ── Layout ─────────────────────────────────────────────────────────────────
pptx.layout = "LAYOUT_WIDE";  // 13.33 x 7.5 inches

// ── Colour palette ─────────────────────────────────────────────────────────
const C = {
  bg:        "#1B2838",   // dark navy – slide background
  bgLight:   "#F7F6F2",   // warm off-white – content slides
  accent:    "#20808D",   // teal
  accentDk:  "#1B474D",   // dark teal
  text:      "#28251D",   // near-black
  textMuted: "#7A7974",
  white:     "#FFFFFF",
  green:     "#2ecc71",
  orange:    "#f39c12",
  red:       "#e74c3c",
};

const W  = 13.33;   // slide width
const H  = 7.5;     // slide height

// ── helper: slide footer (content slides) ──────────────────────────────────
function footer(slide, src) {
  slide.addText(
    src ? [
      { text: "Source: " },
      { text: "Kaggle – Amazon Fine Food Reviews",
        options: { hyperlink: { url: "https://www.kaggle.com/datasets/snap/amazon-fine-food-reviews" } } },
    ] : " ",
    { x: 0.5, y: 7.0, w: 12.3, h: 0.3, fontSize: 9, color: C.textMuted, align: "left" }
  );
}

// ── SLIDE 1 – Title / Team contribution ────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.bg };

  // Top accent bar
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.12, fill: { color: C.accent } });

  // Title
  s.addText("Customer Review Analysis\n& MongoDB Practice", {
    x: 0.7, y: 0.8, w: 8, h: 2.4,
    fontSize: 44, bold: true, color: C.white,
    lineSpacingMultiple: 1.15,
  });

  // Subtitle
  s.addText("IM5211701 Big Data Analytics and Applications  |  Project 1", {
    x: 0.7, y: 3.3, w: 8, h: 0.4,
    fontSize: 14, color: C.accent, italic: true,
  });

  // Divider
  s.addShape(pptx.ShapeType.rect, { x: 0.7, y: 3.85, w: 7.5, h: 0.03, fill: { color: C.accentDk } });

  // Team contribution box
  s.addText("Team Contribution", {
    x: 0.7, y: 4.05, w: 5, h: 0.35,
    fontSize: 13, bold: true, color: C.accent,
  });

  const members = [
    ["Member 1 – [Name]", "Data import, MongoDB schema design, indexing"],
    ["Member 2 – [Name]", "NLP analysis, word frequency, TF-IDF extraction"],
    ["Member 3 – [Name]", "Visualizations, insights, business recommendations"],
    ["Member 4 – [Name]", "Slide preparation, presentation, Q&A coordination"],
  ];
  const rowH = 0.42;
  members.forEach(([name, role], i) => {
    const y = 4.45 + i * rowH;
    s.addText(name, { x: 0.7, y, w: 3.5, h: rowH, fontSize: 12, bold: true, color: C.white });
    s.addText(role, { x: 4.3, y, w: 8.3, h: rowH, fontSize: 12, color: "#B0B8C0" });
  });

  // Bottom accent bar
  s.addShape(pptx.ShapeType.rect, { x: 0, y: H - 0.12, w: W, h: 0.12, fill: { color: C.accent } });
}

// ── SLIDE 2 – Agenda ────────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.bgLight };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.12, fill: { color: C.accent } });

  s.addText("Agenda", {
    x: 0.5, y: 0.35, w: 12, h: 0.7,
    fontSize: 38, bold: true, color: C.accentDk,
  });

  const items = [
    ["01", "Dataset Overview", "Amazon Fine Food Reviews – 568K records, 10+ years"],
    ["02", "MongoDB Import", "Schema design, new fields, indexing strategy"],
    ["03", "NLP Analysis", "Word frequency, TF-IDF, bigrams, text queries"],
    ["04", "Visualisations", "8 charts covering ratings, sentiment, trends"],
    ["05", "Business Insights", "Actionable recommendations for Amazon / marketing"],
  ];

  items.forEach(([num, title, desc], i) => {
    const row = i % 3;
    const col = Math.floor(i / 3);
    const x = 0.5 + col * 6.5;
    const y = 1.35 + row * 1.8;

    // Number chip
    s.addShape(pptx.ShapeType.roundRect, {
      x, y, w: 0.6, h: 0.6,
      fill: { color: C.accent }, rectRadius: 0.08,
    });
    s.addText(num, { x, y, w: 0.6, h: 0.6, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle" });

    s.addText(title, { x: x + 0.75, y, w: 5.5, h: 0.35, fontSize: 15, bold: true, color: C.text });
    s.addText(desc,  { x: x + 0.75, y: y + 0.35, w: 5.5, h: 0.4, fontSize: 11, color: C.textMuted });
  });

  // no footer on agenda
  // footer(s, false);
}

// ── SLIDE 3 – Dataset Overview ──────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.bgLight };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.12, fill: { color: C.accent } });

  s.addText("Dataset Overview", {
    x: 0.5, y: 0.35, w: 12, h: 0.6,
    fontSize: 32, bold: true, color: C.accentDk,
  });

  // KPI cards
  const kpis = [
    ["568,454", "Total Reviews"],
    ["256,059", "Unique Users"],
    ["74,258",  "Products"],
    ["4.18 ★",  "Avg Rating"],
  ];
  kpis.forEach(([val, lbl], i) => {
    const x = 0.5 + i * 3.1;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.15, w: 2.8, h: 1.4, fill: { color: C.accent } });
    s.addText(val, { x, y: 1.22, w: 2.8, h: 0.85, fontSize: 28, bold: true, color: C.white, align: "center" });
    s.addText(lbl, { x, y: 2.1,  w: 2.8, h: 0.4,  fontSize: 11, color: "#CCEDF0", align: "center" });
  });

  // Data dictionary table
  s.addText("Key Fields", { x: 0.5, y: 2.85, w: 4, h: 0.35, fontSize: 14, bold: true, color: C.accentDk });
  const fields = [
    ["ProductId",   "Unique identifier for the product"],
    ["UserId",      "Unique identifier for the reviewer"],
    ["Score",       "Star rating 1–5"],
    ["Time",        "Unix timestamp → converted to ReviewDate"],
    ["Summary",     "Short review headline"],
    ["Text",        "Full review body"],
    ["ProductURL",  "Amazon product link"],
    ["SentimentLabel", "NEW FIELD: Positive / Neutral / Negative"],
    ["HelpfulnessRatio", "NEW FIELD: helpful votes / total votes"],
  ];
  const colW = [2.5, 6.5];
  const hdr = ["Field", "Description"];
  const tblRows = [
    hdr.map(h => ({ text: h, options: { bold: true, color: C.white, fill: C.accentDk, fontSize: 10 } })),
    ...fields.map(([f, d]) => [
      { text: f,  options: { bold: true, color: C.accent, fontSize: 9 } },
      { text: d,  options: { color: C.text, fontSize: 9 } },
    ]),
  ];
  s.addTable(tblRows, {
    x: 0.5, y: 3.25, w: 9.0, colW: [2.5, 6.5],
    border: { type: "solid", color: "#D4D1CA", pt: 0.5 },
    rowH: 0.32,
  });

  // Note about new fields
  s.addText("★  Two new fields added during import: SentimentLabel and HelpfulnessRatio", {
    x: 0.5, y: 6.65, w: 12, h: 0.35,
    fontSize: 10, color: C.accent, italic: true,
  });

  footer(s, true);
}

// ── SLIDE 4 – MongoDB Import & Schema ──────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.bgLight };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.12, fill: { color: C.accent } });

  s.addText("MongoDB Import & Schema Design", {
    x: 0.5, y: 0.35, w: 12, h: 0.6,
    fontSize: 30, bold: true, color: C.accentDk,
  });

  // Left column – approach
  s.addText("Import Approach", { x: 0.5, y: 1.15, w: 5.8, h: 0.35, fontSize: 14, bold: true, color: C.accentDk });
  const steps = [
    "Read Reviews_withURL.csv with pandas",
    "Derive SentimentLabel (Score ≥4 → Positive, =3 → Neutral, ≤2 → Negative)",
    "Derive HelpfulnessRatio = numerator / denominator",
    "Convert Unix timestamp → ReviewDate (datetime)",
    "Batch-insert 10,000 docs/batch into MongoDB",
    "Create indexes: ProductId, UserId, Score, SentimentLabel, ReviewDate",
  ];
  steps.forEach((step, i) => {
    s.addText(`${i + 1}.  ${step}`, {
      x: 0.5, y: 1.6 + i * 0.52, w: 5.8, h: 0.48,
      fontSize: 11, color: C.text,
    });
  });

  // Right column – code block
  s.addText("Python Code Snippet", { x: 6.8, y: 1.15, w: 6, h: 0.35, fontSize: 14, bold: true, color: C.accentDk });
  const code = `from pymongo import MongoClient
import pandas as pd

df = pd.read_csv("Reviews_withURL.csv")

def label(score):
    if score >= 4: return "Positive"
    if score == 3: return "Neutral"
    return "Negative"
df["SentimentLabel"] = df["Score"].apply(label)

df["HelpfulnessRatio"] = df.apply(
    lambda r: r["HelpfulnessNumerator"]
    / r["HelpfulnessDenominator"]
    if r["HelpfulnessDenominator"] > 0
    else None, axis=1)

client = MongoClient("mongodb://localhost:27017")
col = client["amazon_reviews"]["reviews"]
col.insert_many(df.to_dict("records"))
col.create_index("ProductId")
col.create_index("SentimentLabel")`;

  s.addShape(pptx.ShapeType.rect, { x: 6.8, y: 1.5, w: 6.0, h: 5.0, fill: { color: "#1E2A38" } });
  s.addText(code, {
    x: 6.9, y: 1.6, w: 5.8, h: 4.8,
    fontSize: 9.5, fontFace: "Courier New", color: "#9CDCFE",
    lineSpacingMultiple: 1.35,
  });

  footer(s, true);
}

// ── SLIDE 5 – NLP: Word Frequency ──────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.bgLight };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.12, fill: { color: C.accent } });

  s.addText("NLP – Word Frequency Analysis", {
    x: 0.5, y: 0.35, w: 12, h: 0.6,
    fontSize: 30, bold: true, color: C.accentDk,
  });

  // Method description
  s.addText("Method: Python Counter  —  O(n) single-pass tokenisation after stopword removal and HTML stripping.", {
    x: 0.5, y: 1.1, w: 12.3, h: 0.35,
    fontSize: 11, color: C.textMuted, italic: true,
  });

  // Top words table – two columns side by side
  const topWords = [
    ["coffee",  191977], ["tea",       160364], ["love",   155159],
    ["best",    110362], ["little",     87888], ["time",    87136],
    ["price",    87099], ["dog",        83550],
  ];

  s.addText("Top 8 Words Overall", { x: 0.5, y: 1.6, w: 5, h: 0.35, fontSize: 13, bold: true, color: C.accentDk });
  const tblRows2 = [
    [
      { text: "Rank", options: { bold: true, color: C.white, fill: C.accentDk, fontSize: 10 } },
      { text: "Word",  options: { bold: true, color: C.white, fill: C.accentDk, fontSize: 10 } },
      { text: "Count", options: { bold: true, color: C.white, fill: C.accentDk, fontSize: 10 } },
    ],
    ...topWords.map(([w, c], i) => [
      { text: `${i+1}`, options: { color: C.textMuted, fontSize: 10, align: "center" } },
      { text: w,         options: { bold: true, color: C.accent, fontSize: 10 } },
      { text: c.toLocaleString(), options: { color: C.text, fontSize: 10, align: "right" } },
    ]),
  ];
  s.addTable(tblRows2, {
    x: 0.5, y: 2.0, w: 4.8, colW: [0.6, 2.4, 1.8],
    border: { type: "solid", color: "#D4D1CA", pt: 0.5 },
    rowH: 0.36,
  });

  // Bigrams
  s.addText("Top Bigrams (sample of 50K reviews)", { x: 6.0, y: 1.6, w: 6.8, h: 0.35, fontSize: 13, bold: true, color: C.accentDk });
  const bigrams = [
    ["gluten free", 1844], ["peanut butter", 1396], ["highly recommend", 1317],
    ["green tea", 1298],   ["cup coffee", 1111],    ["grocery store", 1076],
    ["subscribe save", 759],["dog loves", 721],      ["dogs love", 717],
    ["love product", 714],
  ];
  const bgramRows = [
    [
      { text: "Bigram", options: { bold: true, color: C.white, fill: C.accentDk, fontSize: 10 } },
      { text: "Count",  options: { bold: true, color: C.white, fill: C.accentDk, fontSize: 10 } },
    ],
    ...bigrams.map(([b, c]) => [
      { text: b, options: { bold: true, color: C.accent, fontSize: 10 } },
      { text: c.toLocaleString(), options: { color: C.text, fontSize: 10, align: "right" } },
    ]),
  ];
  s.addTable(bgramRows, {
    x: 6.0, y: 2.0, w: 6.8, colW: [4.8, 2.0],
    border: { type: "solid", color: "#D4D1CA", pt: 0.5 },
    rowH: 0.35,
  });

  // MongoDB query demo
  s.addText('MongoDB Regex Query – Count reviews containing "delicious":', {
    x: 0.5, y: 5.3, w: 12, h: 0.35, fontSize: 11, bold: true, color: C.accentDk,
  });
  s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 5.7, w: 12.3, h: 0.65, fill: { color: "#1E2A38" } });
  s.addText(
    'col.count_documents({"Text": {"$regex": "delicious", "$options": "i"}})   →   39,114 reviews',
    { x: 0.6, y: 5.75, w: 12.1, h: 0.55, fontSize: 10, fontFace: "Courier New", color: "#9CDCFE" }
  );

  footer(s, true);
}

// ── SLIDE 6 – Chart: Ratings + Sentiment ──────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.bgLight };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.12, fill: { color: C.accent } });

  s.addText("Rating & Sentiment Distribution", {
    x: 0.5, y: 0.35, w: 12, h: 0.6,
    fontSize: 30, bold: true, color: C.accentDk,
  });

  // Embed charts
  s.addImage({ path: "/home/user/workspace/project/figures/fig1_score_distribution.png",
               x: 0.3, y: 1.0, w: 6.3, h: 4.0 });
  s.addImage({ path: "/home/user/workspace/project/figures/fig2_sentiment_pie.png",
               x: 6.8, y: 1.0, w: 6.2, h: 4.0 });

  // Key takeaways
  s.addText("Key Takeaways:", { x: 0.5, y: 5.2, w: 12, h: 0.35, fontSize: 12, bold: true, color: C.accentDk });
  s.addText(
    "78.1% of reviews are Positive (4–5 stars)  •  5-star is the single most common rating (>320K reviews)  •  Only 14.4% Negative",
    { x: 0.5, y: 5.6, w: 12, h: 0.35, fontSize: 11, color: C.text }
  );

  footer(s, true);
}

// ── SLIDE 7 – Chart: Trends over time ─────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.bgLight };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.12, fill: { color: C.accent } });

  s.addText("Review Volume & Rating Trends", {
    x: 0.5, y: 0.35, w: 12, h: 0.6,
    fontSize: 30, bold: true, color: C.accentDk,
  });

  s.addImage({ path: "/home/user/workspace/project/figures/fig3_volume_over_time.png",
               x: 0.3, y: 1.0, w: 6.3, h: 3.8 });
  s.addImage({ path: "/home/user/workspace/project/figures/fig4_avg_score_over_time.png",
               x: 6.8, y: 1.0, w: 6.2, h: 3.8 });

  s.addText("Key Takeaways:", { x: 0.5, y: 5.0, w: 12, h: 0.35, fontSize: 12, bold: true, color: C.accentDk });
  s.addText(
    "Review volume grew ~100× from 2003 to 2012  •  Average score has remained stable near 4.2 despite volume surge  •  Engagement peaked in 2012",
    { x: 0.5, y: 5.4, w: 12, h: 0.45, fontSize: 11, color: C.text }
  );

  footer(s, true);
}

// ── SLIDE 8 – Chart: Word Cloud + Top Words ────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.bgLight };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.12, fill: { color: C.accent } });

  s.addText("NLP Visualisations", {
    x: 0.5, y: 0.35, w: 12, h: 0.6,
    fontSize: 30, bold: true, color: C.accentDk,
  });

  s.addImage({ path: "/home/user/workspace/project/figures/fig6_wordcloud_positive.png",
               x: 0.3, y: 1.0, w: 7.5, h: 3.8 });
  s.addImage({ path: "/home/user/workspace/project/figures/fig5_top20_words.png",
               x: 8.0, y: 1.0, w: 5.0, h: 3.8 });

  s.addText("Key Takeaways:", { x: 0.5, y: 5.0, w: 12, h: 0.35, fontSize: 12, bold: true, color: C.accentDk });
  s.addText(
    "Coffee & tea dominate positive discourse  •  'love', 'best', 'price' signal quality- and value-driven purchases  •  Top bigrams reveal dietary preferences (gluten-free, peanut butter, green tea)",
    { x: 0.5, y: 5.4, w: 12, h: 0.55, fontSize: 11, color: C.text }
  );

  footer(s, true);
}

// ── SLIDE 9 – Chart: Helpfulness + Sentiment Heatmap ──────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.bgLight };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.12, fill: { color: C.accent } });

  s.addText("Helpfulness & Sentiment Over Time", {
    x: 0.5, y: 0.35, w: 12, h: 0.6,
    fontSize: 30, bold: true, color: C.accentDk,
  });

  s.addImage({ path: "/home/user/workspace/project/figures/fig7_helpfulness_by_sentiment.png",
               x: 0.3, y: 1.0, w: 6.3, h: 3.8 });
  s.addImage({ path: "/home/user/workspace/project/figures/fig8_sentiment_heatmap.png",
               x: 6.8, y: 1.0, w: 5.8, h: 3.8 });

  s.addText("Key Takeaways:", { x: 0.5, y: 5.0, w: 12, h: 0.35, fontSize: 12, bold: true, color: C.accentDk });
  s.addText(
    "Negative reviews receive higher helpfulness votes — critical reviews are seen as more useful  •  Positive sentiment share grew from ~70% (2002) to ~80% (2012), suggesting improving product quality or stronger buyer-seller alignment",
    { x: 0.5, y: 5.4, w: 12, h: 0.55, fontSize: 11, color: C.text }
  );

  footer(s, true);
}

// ── SLIDE 10 – Business Recommendations ───────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.bgLight };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.12, fill: { color: C.accent } });

  s.addText("Business Recommendations", {
    x: 0.5, y: 0.35, w: 12, h: 0.6,
    fontSize: 30, bold: true, color: C.accentDk,
  });

  const recs = [
    {
      title: "Highlight Helpful Negative Reviews",
      body:  "Critical reviews receive disproportionately high helpfulness votes. Amazon should surface high-ratio negative reviews prominently to improve buyer trust and reduce returns.",
      tag: "For Amazon Product Team",
    },
    {
      title: "Target Dietary & Lifestyle Niches",
      body:  "'Gluten free', 'peanut butter', 'green tea', and 'dog loves' are top bigrams. Marketing teams should develop dedicated landing pages and ad campaigns for these high-engagement niches.",
      tag: "For Marketing Teams",
    },
    {
      title: "Leverage 'Subscribe & Save' Language",
      body:  "'Subscribe save' ranks in top bigrams. Promote subscription bundles more aggressively in the food category — customers already associate value with this programme.",
      tag: "For Amazon Subscription Team",
    },
    {
      title: "Monitor Sentiment Trends as Quality Signal",
      body:  "Positive sentiment grew from 70% (2002) to 80% (2012). Use monthly sentiment tracking in MongoDB as an early-warning signal for product or supplier quality deterioration.",
      tag: "For Category Managers",
    },
  ];

  recs.forEach((r, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.5 + col * 6.4;
    const y = 1.3 + row * 2.6;
    s.addShape(pptx.ShapeType.rect, { x, y, w: 6.0, h: 2.3, fill: { color: "#EEF9FA" }, line: { color: C.accent, pt: 0.8 } });
    s.addShape(pptx.ShapeType.rect, { x, y, w: 6.0, h: 0.38, fill: { color: C.accent } });
    s.addText(r.title, { x: x + 0.12, y: y + 0.03, w: 5.76, h: 0.35, fontSize: 11, bold: true, color: C.white });
    s.addText(r.body,  { x: x + 0.15, y: y + 0.48, w: 5.7,  h: 1.4,  fontSize: 10, color: C.text });
    s.addText(r.tag,   { x: x + 0.15, y: y + 1.9,  w: 5.7,  h: 0.3,  fontSize: 9, color: C.accent, italic: true });
  });

  footer(s, true);
}

// ── SLIDE 11 – Conclusion ──────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.12, fill: { color: C.accent } });

  s.addText("Conclusion", {
    x: 0.7, y: 0.7, w: 11, h: 0.7,
    fontSize: 42, bold: true, color: C.white,
  });

  const bullets = [
    "Successfully imported 568,454 reviews into MongoDB with 2 new engineered fields",
    "Demonstrated efficient NLP: Counter-based O(n) word frequency, TF-IDF, and MongoDB text-index queries",
    "8 visualisations uncovered rating distribution, temporal growth, topical bigrams, and helpfulness patterns",
    "Actionable insights delivered for Amazon's product, marketing, and subscription teams",
  ];

  bullets.forEach((b, i) => {
    s.addShape(pptx.ShapeType.rect, { x: 0.7, y: 1.65 + i * 1.1, w: 0.07, h: 0.6, fill: { color: C.accent } });
    s.addText(b, { x: 1.0, y: 1.65 + i * 1.1, w: 11.8, h: 0.6, fontSize: 14, color: C.white });
  });

  s.addText("Thank you  |  Q&A", {
    x: 0.7, y: 6.6, w: 11, h: 0.5,
    fontSize: 16, bold: true, color: C.accent, align: "center",
  });

  s.addShape(pptx.ShapeType.rect, { x: 0, y: H - 0.12, w: W, h: 0.12, fill: { color: C.accent } });
}

// ── Write file ─────────────────────────────────────────────────────────────
const outPath = "/home/user/workspace/project/CustomerReviewAnalysis.pptx";
pptx.writeFile({ fileName: outPath }).then(() => {
  console.log("PPTX saved to", outPath);
}).catch(err => {
  console.error("Error:", err);
  process.exit(1);
});
