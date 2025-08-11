from rapidfuzz import process, fuzz
import sys

def _prompt_yes_no(question: str) -> str:
    # GUI watches for [[PROMPT:YN]] lines
    print(f"[[PROMPT:YN]] {question}", flush=True)
    return input().strip().lower()

def _prompt_text(question: str) -> str:
    # GUI watches for [[PROMPT:TEXT]] lines
    print(f"[[PROMPT:TEXT]] {question}", flush=True)
    return input().strip()

# ─────────────── 4) MATCH & DEBUG ───────────────
def match_entries_debug(entries, db, threshold=80):
    """Return [(raw_text, amount, db_row)] with verbose logs."""
    keywords = db["E"].astype(str).str.strip().tolist()
    matches  = []

    for raw_txt, amt in entries:
        print("\nBANK ROW")
        print(f"   Text   : {raw_txt!r}")
        print(f"   Amount : {amt}")

        # 4-a) exact substring in db["E"]
        clean   = raw_txt.replace(" ", "")
        subset  = db[db["E"].apply(lambda k: str(k).replace(' ', '') in clean)]
        if not subset.empty:
            hit = subset.iloc[0]
            print("   Exact match:")
            print(f"      Keyword     : {hit['E']!r}")
            print(f"      Customer ID : {hit['F']}  Clean Name : {hit['G']!r}")
            matches.append((raw_txt, amt, hit))
            continue

        # 4-b) fuzzy fallback
        best = process.extractOne(
            clean, keywords, scorer=fuzz.partial_ratio
        )
        if best:
            best_kw, score, _ = best
            print(f"   Fuzzy best : {best_kw!r}  (score {score:.1f})")
            if score >= threshold:
                idx = keywords.index(best_kw)
                hit = db.iloc[idx]
                print("   Accepted fuzzy match")
                matches.append((raw_txt, amt, hit))
                continue
            else:
                print(f"   Score {score:.1f} < threshold {threshold}")
        else:
            print("    No fuzzy candidate at all")

    print(f"\n🔗 Matched {len(matches)}/{len(entries)} rows")
    return matches


def match_entries_interactive(entries, db, threshold=80):
    """
    entries: list of (raw_txt, amt)
    db: DataFrame with columns E (keyword), F (cust_id), G (clean_name)
    """

    keywords = db["E"].astype(str).str.strip().tolist()
    matches = []
    skipped  = []
    
    # 1) filter out zero‐amounts
    entries = [(txt, amt) for txt, amt in entries if amt and float(amt) != 0]

    for raw_txt, amt in entries:
        print("\nROW:")
        print(f"  desc  = {raw_txt!r}")
        print(f"  amount= {amt}")

        # 2) exact substring
        key_clean = raw_txt.replace(" ", "")
        subset = db[ db["E"].apply(lambda k: key_clean.find(str(k).replace(" ","")) >= 0) ]
        if not subset.empty:
            hit = subset.iloc[0]
            print("  Exact match:")
            print(f"     → {hit['E']!r}  [{hit['F']}] {hit['G']}")
            matches.append((raw_txt, amt, hit))
            continue

        # 3) fuzzy fallback
        best, score, _ = process.extractOne(
            key_clean, keywords, scorer=fuzz.partial_ratio
        )
        print(f"  Best fuzzy: {best!r}  (score {score:.1f})")
        idx = keywords.index(best)
        hit = db.iloc[idx]

        # 4) ask user
        # ans = input(f"    接受 (y/n) ").strip().lower()
        ans = _prompt_yes_no("接受這個配對嗎？(y/n)")
        if ans in ("", "y", "yes"):
            matches.append((raw_txt, amt, hit))
        else:
            # manual override
            # manual = input("    請輸入客戶ID（或留空以跳過）：").strip()
            manual = _prompt_text("請輸入客戶ID（或留空以跳過）：")
            if manual:
                # look up manual ID in db
                row = db[ db["F"].astype(str) == manual ]
                if not row.empty:
                    hit2 = row.iloc[0]
                    matches.append((raw_txt, amt, hit2))
                else:
                    print(f"    ID {manual!r} not found—skipping.")
                    skipped.append((raw_txt, amt))
            else:
                print("    skipped.")
                skipped.append((raw_txt, amt))

    print(f"Done: {len(matches)} matched, {len(skipped)} skipped")
    return matches, skipped
