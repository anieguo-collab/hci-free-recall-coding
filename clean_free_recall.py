"""
Clean and score free-recall responses (Swipe vs Grid study).

Pipeline
    1. Normalize   : lowercase, strip accents/punctuation, collapse spaces
    2. Split       : one cell -> separate answers (only ITEM_SEPARATOR is used)
    3. Spell check : dictionary lookup, plural/synonym rules, fuzzy matching
    4. Learn       : frequent unknown words are promoted into the dictionary
    5. Match       : each cleaned item is fuzzy-matched to a product in the catalog
    6. Report      : recall-by-product sheet, cleaned text, matched IDs, corrections,
                     learned words, review list

Usage
    python clean_free_recall.py                      (reads HCI2000_..._Dummy_Data.xlsx next to this script)
"""
import re
import sys
from pathlib import Path
import unicodedata
from collections import Counter, defaultdict

import pandas as pd
from openpyxl.styles import Alignment

# ----------------------------------------------------------------------------
# CONFIG - edit these values to tune the behaviour
# ----------------------------------------------------------------------------
# Input workbook: read directly from the same folder as this script
DEFAULT_INPUT = Path(__file__).with_name("HCI2000_Swipe_vs_Grid_40_Participants_Dummy_Data.xlsx")
DEFAULT_OUTPUT = Path(__file__).with_name("cleaned_free_recall.xlsx")

SHEET_SESSIONS = "Session Results"
SHEET_CATALOG = "Product Catalog"
HEADER_ROW = 3                      # zero-based row index of the header line

COL_PID = "Participant ID"
COL_COND = "Condition #"
COL_RECALL = "Free recall – verbatim (Form)"
COL_VIEWED = "Viewed product IDs (unique)"
COL_IFACE = "Interface"             # optional: used to split the product table by interface
COL_MANUAL = "Scored recall IDs"    # optional: manual scores used for validation

# Words that carry no product information (never spell-checked, ignored when matching)
STOPWORDS = {"a", "an", "the", "of", "with", "and", "in", "it", "i", "my",
             "some", "one", "that", "this", "not", "no", "was", "had"}

# Answers that mean "I remember nothing"
NON_ANSWERS = {"none", "nothing", "idk", "dont know", "dont remember", "na",
               "n a", "no answer", "cant remember", "not sure", "x"}

# Manual synonym map: participant word -> catalog word (extend as needed)
SYNONYMS = {"trouser": "trousers", "jumper": "sweater", "tshirt": "tee",
            "tees": "tee", "denims": "denim"}

# Extra words that are valid but not in any product name
EXTRA_WORDS = set()

# The form collects answers one at a time, and the workbook stores one session's answers
# joined by this separator inside a single cell. Set to None if every cell already holds
# exactly one answer (each cell is then treated as one item, nothing is split).
ITEM_SEPARATOR = ";"

MIN_FREQ = 3             # an unknown word must occur at least this many times ...
MIN_PARTICIPANTS = 2     # ... and be used by at least this many participants
MATCH_MIN = 0.60         # minimum token overlap (Jaccard) to accept a product match
MATCH_MARGIN = 0.10      # best match must beat the runner-up by this margin


# ----------------------------------------------------------------------------
# 1. Text normalization
# ----------------------------------------------------------------------------
def strip_accents(text):
    """Remove accents, e.g. 'crème' -> 'creme'."""
    decomposed = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def split_items(raw):
    """Un-join the answers stored in one cell. Only ITEM_SEPARATOR is used, so words
    such as 'and' or characters such as ',' inside a product description stay intact."""
    text = str(raw)
    if ITEM_SEPARATOR is None:
        return [text] if text.strip() else []
    return [part for part in text.split(ITEM_SEPARATOR) if part.strip()]


def normalize_item(item):
    """Lowercase, remove accents and punctuation, collapse whitespace."""
    item = strip_accents(unicodedata.normalize("NFKC", item).lower())
    item = item.replace("\u00a0", " ").replace("'", "")
    item = re.sub(r"[-_]", " ", item)            # 'wide-leg' -> 'wide leg'
    item = re.sub(r"[^a-z0-9\s]", " ", item)     # drop remaining punctuation
    return " ".join(item.split())


def tokenize(item):
    return item.split()


# ----------------------------------------------------------------------------
# 2. Dictionary and spell checking
# ----------------------------------------------------------------------------
def build_dictionary(catalog):
    """Dictionary = every word that appears in a catalog product name + extras."""
    words = set(EXTRA_WORDS)
    for name in catalog["Product name"].dropna():
        words.update(tokenize(normalize_item(name)))
    return words


def edit_distance(a, b):
    """Damerau-Levenshtein (optimal string alignment): counts insert, delete,
    substitute and adjacent swap ('skrit' -> 'skirt' = 1) as one edit each."""
    la, lb = len(a), len(b)
    d = [[0] * (lb + 1) for _ in range(la + 1)]
    for i in range(la + 1):
        d[i][0] = i
    for j in range(lb + 1):
        d[0][j] = j
    for i in range(1, la + 1):
        for j in range(1, lb + 1):
            cost = 0 if a[i - 1] == b[j - 1] else 1
            d[i][j] = min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + cost)
            if i > 1 and j > 1 and a[i - 1] == b[j - 2] and a[i - 2] == b[j - 1]:
                d[i][j] = min(d[i][j], d[i - 2][j - 2] + 1)
    return d[la][lb]


def max_edits(word):
    """Short words are too risky to auto-correct, long words tolerate more typos."""
    n = len(word)
    return 0 if n <= 3 else (1 if n <= 6 else 2)


def correct_token(tok, dictionary):
    """Return (corrected_token, rule, distance). Rule is one of:
    ok / synonym / plural / fuzzy / ambiguous / unknown."""
    if tok in dictionary or tok in STOPWORDS or tok.isdigit():
        return tok, "ok", 0
    if tok in SYNONYMS and SYNONYMS[tok] in dictionary:
        return SYNONYMS[tok], "synonym", 0
    for suffix in ("es", "s"):                       # simple plural rule
        if tok.endswith(suffix) and tok[: -len(suffix)] in dictionary:
            return tok[: -len(suffix)], "plural", 1
    limit = max_edits(tok)
    if limit == 0:
        return tok, "unknown", 0
    scored = sorted((edit_distance(tok, w), w) for w in dictionary if abs(len(w) - len(tok)) <= limit)
    scored = [(dist, w) for dist, w in scored if dist <= limit]
    if not scored:
        return tok, "unknown", 0
    best = scored[0][0]
    winners = [w for dist, w in scored if dist == best]
    if len(winners) > 1:                             # tie -> do not guess
        return tok, "ambiguous", best
    return winners[0], "fuzzy", best


# ----------------------------------------------------------------------------
# 3. Product matching
# ----------------------------------------------------------------------------
def build_product_tokens(catalog):
    """Product ID -> sorted list of content words in its name (stopwords removed)."""
    return {row["Product ID"]: sorted(set(tokenize(normalize_item(row["Product name"]))) - STOPWORDS)
            for _, row in catalog.iterrows()}


def words_similar(a, b):
    """Typo-tolerant word comparison used at product level. The allowed number of edits
    grows with word length; the other words in the item give extra context, so this can
    be a little more tolerant than the word-level spell checker."""
    if a == b:
        return True
    n = max(len(a), len(b))
    tolerance = 0 if n <= 2 else (1 if n <= 5 else 2)
    return abs(len(a) - len(b)) <= tolerance and edit_distance(a, b) <= tolerance


def count_pairs(content, product_words):
    """Greedily pair each item word with its closest unused product word."""
    used, pairs = set(), 0
    for tok in content:
        best = None
        for word in product_words:
            if word not in used and words_similar(tok, word):
                dist = edit_distance(tok, word)
                if best is None or dist < best[0]:
                    best = (dist, word)
        if best:
            used.add(best[1])
            pairs += 1
    return pairs


def match_item(tokens, candidates, product_tokens):
    """Match an item to the catalog with a fuzzy Jaccard score
    (paired words / all distinct words). Returns (product_id or None, score, status)."""
    content = sorted({t for t in tokens if t not in STOPWORDS})
    if not content:
        return None, 0.0, "empty"
    scores = []
    for pid in candidates:
        pairs = count_pairs(content, product_tokens[pid])
        scores.append((pairs / (len(content) + len(product_tokens[pid]) - pairs), pid))
    scores.sort(reverse=True)
    if not scores or scores[0][0] < MATCH_MIN:
        return None, scores[0][0] if scores else 0.0, "unmatched"
    runner_up = scores[1][0] if len(scores) > 1 else 0.0
    if scores[0][0] - runner_up < MATCH_MARGIN:       # several products fit equally well
        return None, scores[0][0], "ambiguous"
    return scores[0][1], scores[0][0], "matched"


# ----------------------------------------------------------------------------
# 4. Main pipeline
# ----------------------------------------------------------------------------
def clean_free_recall(sessions, catalog):
    dictionary = build_dictionary(catalog)
    product_tokens = build_product_tokens(catalog)

    # --- Step A: normalize + split every response -------------------------
    records = []                                     # one dict per recalled item
    for idx, row in sessions.iterrows():
        raw_response = row[COL_RECALL]
        if pd.isna(raw_response) or not str(raw_response).strip():
            continue
        for raw_item in split_items(raw_response):
            norm = normalize_item(raw_item)
            records.append({"row": idx, "pid": row[COL_PID], "cond": row[COL_COND],
                            "raw_item": raw_item.strip(), "norm": norm,
                            "tokens": tokenize(norm)})

    # --- Step B: spell-check every unique token once ------------------------
    token_freq, token_people = Counter(), defaultdict(set)
    for rec in records:
        for tok in rec["tokens"]:
            token_freq[tok] += 1
            token_people[tok].add(rec["pid"])
    corrections = {tok: correct_token(tok, dictionary) for tok in token_freq}

    # --- Step C: learn frequent unknown words ---------------------------------
    learned = []
    for tok, (_, rule, _) in corrections.items():
        if rule in ("unknown", "ambiguous") and token_freq[tok] >= MIN_FREQ \
                and len(token_people[tok]) >= MIN_PARTICIPANTS:
            dictionary.add(tok)
            learned.append({"word": tok, "frequency": token_freq[tok],
                            "participants": len(token_people[tok])})
    # second pass: remaining unknown words may be typos of a newly learned word
    for tok, (_, rule, _) in list(corrections.items()):
        if rule in ("unknown", "ambiguous"):
            corrections[tok] = correct_token(tok, dictionary) if tok not in dictionary else (tok, "learned", 0)

    # --- Step D: rebuild cleaned items and match to products ----------------
    session_ids = {}
    if COL_VIEWED in sessions.columns:
        for idx, row in sessions.iterrows():
            session_ids[idx] = [x.strip().upper() for x in re.split(r"[;,]", str(row[COL_VIEWED])) if x.strip()]
    all_ids = list(product_tokens)

    item_rows = []
    for rec in records:
        fixed = [corrections[t][0] for t in rec["tokens"]]
        cleaned = " ".join(fixed)
        if cleaned in NON_ANSWERS or not cleaned:
            pid_, score, status = None, 0.0, "non_answer"
        else:
            candidates = [p for p in session_ids.get(rec["row"], all_ids) if p in product_tokens] or all_ids
            pid_, score, status = match_item(fixed, candidates, product_tokens)
        item_rows.append({"Participant ID": rec["pid"], "Condition #": rec["cond"],
                          "raw_item": rec["raw_item"], "cleaned_item": cleaned,
                          "matched_product_id": pid_, "match_score": round(score, 2),
                          "status": status, "_row": rec["row"]})
    items = pd.DataFrame(item_rows)

    # --- Step E: roll items back up to one row per session ------------------
    keep = [COL_PID, COL_COND, COL_RECALL] + ([COL_IFACE] if COL_IFACE in sessions.columns else [])
    out = sessions[keep].copy()
    out["recall_clean"] = ""
    out["recall_ids_auto"] = ""
    out["n_items"] = 0
    out["n_matched"] = 0
    out["n_need_review"] = 0
    for idx, group in items.groupby("_row"):
        ids = list(dict.fromkeys(group["matched_product_id"].dropna()))     # unique, keep order
        out.at[idx, "recall_clean"] = "; ".join(group["cleaned_item"])
        out.at[idx, "recall_ids_auto"] = ", ".join(ids)
        out.at[idx, "n_items"] = len(group)
        out.at[idx, "n_matched"] = len(ids)
        out.at[idx, "n_need_review"] = int(group["status"].isin(["unmatched", "ambiguous"]).sum())
    if COL_MANUAL in sessions.columns:               # compare with manual scoring if present
        manual = sessions[COL_MANUAL].fillna("").map(lambda s: {x.strip().upper() for x in str(s).split(",") if x.strip()})
        auto = out["recall_ids_auto"].map(lambda s: {x.strip() for x in s.split(",") if x.strip()})
        out["agrees_with_manual"] = manual == auto

    # --- Step F: reports -------------------------------------------------------
    corr_rows = [{"original_word": t, "corrected_to": c, "rule": r, "edit_distance": d,
                  "frequency": token_freq[t]}
                 for t, (c, r, d) in corrections.items() if r != "ok"]
    corr = pd.DataFrame(corr_rows).sort_values("frequency", ascending=False) if corr_rows else pd.DataFrame()
    learned_df = pd.DataFrame(learned)
    review = items[items["status"].isin(["unmatched", "ambiguous"])].drop(columns="_row")
    return out, items.drop(columns="_row"), corr, learned_df, review


# ----------------------------------------------------------------------------
# 5. Product-centred view: which participants remembered each product
# ----------------------------------------------------------------------------
def split_ids(text):
    """'S1-03, S1-07' -> ['S1-03', 'S1-07'] (empty or NaN -> [])."""
    if pd.isna(text):
        return []
    return [x.strip().upper() for x in re.split(r"[;,]", str(text)) if x.strip()]


def recall_by_product(out, sessions, catalog):
    """One row per product with the participants who recalled it (overall and per interface).
    `out` comes from clean_free_recall(); its index matches `sessions`."""
    has_iface = COL_IFACE in out.columns
    recalled, viewed = defaultdict(list), defaultdict(list)     # product ID -> [(participant, interface)]
    for idx, row in out.iterrows():
        iface = str(row[COL_IFACE]).strip().title() if has_iface else ""
        for prod in split_ids(row["recall_ids_auto"]):
            recalled[prod].append((row[COL_PID], iface))
        for prod in split_ids(sessions.at[idx, COL_VIEWED]) if COL_VIEWED in sessions.columns else []:
            viewed[prod].append((row[COL_PID], iface))

    rows = []
    for _, prod in catalog.iterrows():
        pid = prod["Product ID"]
        who = sorted({p for p, _ in recalled[pid]})
        row = {"Product ID": pid, "Product set": prod["Product set"], "Product name": prod["Product name"],
               "Participants who recalled": ", ".join(who), "# recalled": len(who),
               "# viewed": len({p for p, _ in viewed[pid]}) or None}
        row["Recall rate"] = round(len(who) / row["# viewed"], 3) if row["# viewed"] else None
        if has_iface:
            for iface in ("Swipe", "Grid"):
                names = sorted({p for p, i in recalled[pid] if i == iface})
                row[f"Recalled under {iface}"] = ", ".join(names)
                row[f"# {iface}"] = len(names)
        rows.append(row)
    return pd.DataFrame(rows)


# ----------------------------------------------------------------------------
# 6. Command-line entry point
# ----------------------------------------------------------------------------
def main(src=DEFAULT_INPUT, dst=DEFAULT_OUTPUT):
    sessions = pd.read_excel(src, sheet_name=SHEET_SESSIONS, header=HEADER_ROW)
    catalog = pd.read_excel(src, sheet_name=SHEET_CATALOG, header=HEADER_ROW)
    sessions = sessions.dropna(subset=[COL_PID]).reset_index(drop=True)
    out, items, corr, learned, review = clean_free_recall(sessions, catalog)
    by_product = recall_by_product(out, sessions, catalog)
    with pd.ExcelWriter(dst, engine="openpyxl") as writer:
        by_product.to_excel(writer, sheet_name="Recall by product", index=False)
        out.to_excel(writer, sheet_name="Cleaned sessions", index=False)
        items.to_excel(writer, sheet_name="Items", index=False)
        corr.to_excel(writer, sheet_name="Corrections", index=False)
        learned.to_excel(writer, sheet_name="Learned words", index=False)
        review.to_excel(writer, sheet_name="Needs review", index=False)
        for ws in writer.sheets.values():           # simple auto column width
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col[:200]) + 2
                ws.column_dimensions[col[0].column_letter].width = min(width, 60)
        ws_prod = writer.sheets["Recall by product"]     # wrap long participant lists, show rate as %
        for row in ws_prod.iter_rows(min_row=2):
            for cell in row:
                header = ws_prod.cell(1, cell.column).value
                if str(header).startswith(("Participants", "Recalled under")):
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                if header == "Recall rate":
                    cell.number_format = "0%"
    print(f"items: {len(items)} | corrections: {len(corr)} | learned words: {len(learned)} "
          f"| need review: {len(review)}")


if __name__ == "__main__":
    # No arguments -> read DEFAULT_INPUT and write DEFAULT_OUTPUT.
    main(*sys.argv[1:3])
