"""
Sentence Splitting Inspector - Sinhala vs Tamil (single file pair)
=====================================================================

Purpose: JUST check how well the sentence splitter is working, before
trusting it for alignment. No embeddings, no alignment, no similarity -
this only runs the tokenizers and shows you the results side by side so
you can visually spot bad splits (abbreviations, numbered clauses,
initials, etc. wrongly cut).

Input : the exact path to one Sinhala .txt file and one Tamil .txt file
        (edit SINHALA_FILE / TAMIL_FILE below, or pass --sinhala_file /
        --tamil_file on the command line).

Output: one .xlsx workbook with a single sheet, columns:
        idx | si_id | si_sentence | ta_id | ta_sentence
        Rows are placed side by side by POSITION ONLY (row 1 = 1st
        Sinhala sentence next to 1st Tamil sentence, etc.) - this is NOT
        an alignment, just a positional view for eyeballing whether the
        two sentence counts and boundaries look reasonable.

Requirements:
    pip install sinling tamil-tokenizer openpyxl pandas --break-system-packages
"""

import os
import re
import pandas as pd

# ---------------------------------------------------------------------------
# INPUT / OUTPUT PATHS - edit these and just run the script directly.
# ---------------------------------------------------------------------------
SINHALA_FILE = os.path.join(os.path.dirname(__file__), "input", "2008-38 s.txt")
TAMIL_FILE = os.path.join(os.path.dirname(__file__), "input", "2008-38 t.txt")

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")
OUTPUT_FILENAME = "sentence_split_check-ACT.xlsx"

# Use the wtpsplit / SaT model as the primary sentence splitter instead of
# sinling / tamil-tokenizer. SaT is a single multilingual model (no per-
# language code needed) trained to be robust to missing/irregular
# punctuation - exactly the kind of noisy OCR text we're dealing with.
# Falls back to sinling/tamil-tokenizer/naive automatically if the model
# can't be loaded (e.g. not installed).
USE_WTPSPLIT = False
SAT_MODEL_NAME = "sat-12l-sm"   # good speed/accuracy tradeoff; try "sat-12l-sm" for max accuracy

# When True (default): apply our custom pre/post-processing on top of
# whichever tokenizer runs - numbered-marker pre-chunking, colon/abbreviation
# masking, short-fragment merging.
# When False: run the chosen tokenizer (SaT, if USE_WTPSPLIT) on the raw
# paragraph text with NONE of our custom rules - useful to see exactly how
# good SaT is on its own, without anything compensating for its mistakes.
USE_CUSTOM_RULES = True

# ---------------------------------------------------------------------------
# TEXT LOADING + PARAGRAPH SPLITTING
# ---------------------------------------------------------------------------

def load_text(path: str) -> str:
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


# A line made only of dashes is layout (a heading underline), not content.
# It carries no sentence punctuation, so no tokenizer will break on it - left
# in place it silently glues the heading above it to the body below.
RULE_LINE_RE = re.compile(r"^[-–—_=\s|]*[-–—_=]{4,}[-–—_=\s|]*$")


def split_paragraph_segments(text: str):
    """Returns a list of paragraphs, each a LIST OF SEGMENTS. A dashes-only
    line is dropped as layout, and the line immediately above it is a heading,
    so it becomes a segment of its own - this is what stops headings from
    being swallowed into the following body sentence. These documents never
    put a full stop after a heading, so the line break is the ONLY boundary
    signal available; joining lines with a space would destroy it."""
    paras = []
    for raw in re.split(r"\n\s*\n", text.strip()):
        lines = [l.strip() for l in raw.split("\n") if l.strip()]
        if not lines:
            continue
        segments, buf = [], []
        for line in lines:
            if RULE_LINE_RE.match(line):
                if buf:
                    segments.append(" ".join(buf[:-1]).strip())  # body before
                    segments.append(buf[-1].strip())             # the heading
                    buf = []
                continue
            buf.append(line)
        if buf:
            segments.append(" ".join(buf).strip())
        segments = [s for s in segments if s]
        if segments:
            paras.append(segments)
    return paras


# ---------------------------------------------------------------------------
# SENTENCE SPLITTING (same logic as the main pipeline)
# ---------------------------------------------------------------------------

SINHALA_ABBREVIATIONS = ["මහ", "ලංකා", "අංක", "නො"]   # placeholders - fill in from your corpus scan
TAMIL_ABBREVIATIONS = ["எண்", "இல", "மா"]              # placeholders - fill in from your corpus scan

DECIMAL_RE = re.compile(r"(?<=\d)\.(?=\d)")
URL_EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+|https?://\S+|www\.\S+")

# ---- Fix for PROBLEM 1: numbered-list clauses (e.g. "... කිරීම 02. ආයතන...")
# getting merged into one giant sentence because there is no real sentence-
# ending punctuation between clauses in the OCR text, just a bare list
# marker. We FORCE a boundary right before a 1-2 digit number + "." that is
# followed by whitespace, wherever it appears (not just at the very start),
# since that is almost always the start of a new numbered clause in these
# government-document formats. We split the raw text into chunks at these
# markers BEFORE handing anything to the tokenizer, so the tokenizer can
# never re-merge them (it only ever sees one chunk at a time).
#
# These documents also use single-letter markers for lettered sub-points -
# either the classical vowel ordering (අ./ආ./ඇ./ඈ./ඉ./ඊ.../ in Sinhala,
# அ./ஆ./இ./ஈ./உ./ஊ.../ in Tamil) or plain Roman letters (a./b./c.../ or
# A./B./C.../) - so they all get the same treatment. Restricted to a SINGLE
# letter character (not 1-2 chars like the digit markers) so we never
# accidentally match an ordinary short word ending a sentence (e.g. Sinhala
# "ඇත." = "is/exists", which is 2 characters + ".").
#
# On top of that, points can also be numbered with Roman numerals
# (I./II./III./IV./.../XX. or lowercase i./ii./iii./...), so those get
# matched too - up to 8 letters, enough for numerals into the low
# thousands (e.g. "VIII.", "XVIII.").
_SINHALA_LETTER_MARKERS = "අආඇඈඉඊඋඌඍඎඏඐඑඒඓඔඕඖ"
_TAMIL_LETTER_MARKERS = "அஆஇஈஉஊஎஏஐஒஓஔ"
_ROMAN_LETTER_MARKERS = "a-zA-Z"
_ROMAN_NUMERAL_RE_PART = r"[IVXLCDM]{1,8}|[ivxlcdm]{1,8}"
NUMBERED_LIST_MARKER_RE = re.compile(
    rf"(?<!\S)(\d{{1,2}}\.|(?:{_ROMAN_NUMERAL_RE_PART})\.|[{_SINHALA_LETTER_MARKERS}{_TAMIL_LETTER_MARKERS}{_ROMAN_LETTER_MARKERS}]\.)(?=\s)"
)


def split_on_numbered_markers(text: str):
    """Split text into (marker, rest) pairs at numbered-list markers (e.g.
    '02.'). marker is None for the first chunk if the text doesn't start
    with a marker. Keeping the marker separate from the text handed to the
    tokenizer means the tokenizer can never swallow/strip its trailing dot
    (sinling/tamil-tokenizer both drop the "." when they see a bare number
    like "1." as a token) - we reattach the marker verbatim afterwards."""
    parts = NUMBERED_LIST_MARKER_RE.split(text)
    if len(parts) == 1:
        return [(None, text)]
    chunks = []
    first = parts[0].strip()
    if first:
        chunks.append((None, first))
    i = 1
    while i < len(parts):
        marker = parts[i]
        rest = parts[i + 1] if i + 1 < len(parts) else ""
        rest = rest.strip()
        if marker or rest:
            chunks.append((marker, rest))
        i += 2
    return chunks


# ---- Fix for PROBLEM 2 & 3: the tokenizer over-splitting a bare list
# number ("01", "02"...) or a short initial ("ඡේ", "ஜே") off into its own
# standalone "sentence". We merge any fragment that is too short or looks
# like a bare number/initial FORWARD into the next fragment, so these
# reattach to the clause/name they actually belong to, instead of floating
# as their own meaningless one-token "sentence".
MIN_SENTENCE_LEN =15       # fragments shorter than this get merged forward
BARE_NUMBER_RE = re.compile(r"^\d{1,3}\.?$")


def merge_short_fragments(sentences, min_len: int = MIN_SENTENCE_LEN):
    merged = []
    buffer = ""
    for s in sentences:
        buffer = f"{buffer} {s}".strip() if buffer else s
        stripped = buffer.strip()
        if BARE_NUMBER_RE.match(stripped) or len(stripped) < min_len:
            continue  # too short / just a number - keep accumulating, don't flush yet
        merged.append(stripped)
        buffer = ""
    if buffer.strip():
        # leftover short fragment at the end - attach to the previous
        # sentence if there is one, otherwise keep it as-is
        if merged:
            merged[-1] = f"{merged[-1]} {buffer.strip()}"
        else:
            merged.append(buffer.strip())
    return merged


def _protect_false_boundaries(text: str, abbreviations):
    protected = text
    protected = DECIMAL_RE.sub("<DEC>", protected)
    for match in URL_EMAIL_RE.finditer(protected):
        protected = protected.replace(match.group(0), match.group(0).replace(".", "<DOT>"))
    # short parenthesised abbreviations like "(පා.ම.)" / "(பா.உ.)" (both mean
    # "M.P.") get their dots eaten by the tokenizer the same way bare list
    # markers do. Mask every "." inside a (...) span so the tokenizer never
    # treats it as a sentence boundary - a period rarely legitimately ends a
    # sentence while still inside unclosed parentheses in these documents.
    protected = re.sub(r"\([^()]*\)", lambda m: m.group(0).replace(".", "<PDOT>"), protected)
    for abbr in abbreviations:
        protected = re.sub(rf"(?<={re.escape(abbr)})\.", "<ABR>", protected)
    # colons in these documents are almost always label:value separators
    # (e.g. "චක්‍රලේඛ: 02/2016", "අංකය: HAF-1/09"), never real sentence
    # endings, but some tokenizers (sinling in particular) treat ":" as a
    # sentence boundary anyway. Mask it so the tokenizer never sees it.
    protected = protected.replace(":", "<COL>")
    return protected


def _restore(text: str) -> str:
    return (
        text.replace("<DEC>", ".")
        .replace("<DOT>", ".")
        .replace("<PDOT>", ".")
        .replace("<ABR>", ".")
        .replace("<COL>", ":")
    )


def naive_sentence_split(text: str, abbreviations):
    protected = _protect_false_boundaries(text, abbreviations)
    pieces = re.split(r"(?<=[.?!])\s+", protected)
    return [_restore(p).strip() for p in pieces if p.strip()]


# ---- wtpsplit / SaT loader (single multilingual model, no language code
# needed). Loaded once and reused across all calls.
_SAT_MODEL = None


def _get_sat_model():
    global _SAT_MODEL
    if _SAT_MODEL is None:
        from wtpsplit import SaT
        _SAT_MODEL = SaT(SAT_MODEL_NAME)
    return _SAT_MODEL


def _tokenize_chunk_sat(text: str, abbreviations):
    protected = _protect_false_boundaries(text, abbreviations)
    sat = _get_sat_model()
    sents = sat.split(protected)
    return [_restore(s).strip() for s in sents if s.strip()]


def _tokenize_chunk_sinhala(text: str):
    if USE_WTPSPLIT:
        try:
            return _tokenize_chunk_sat(text, SINHALA_ABBREVIATIONS)
        except Exception:
            pass  # fall through to sinling / naive below
    try:
        from sinling import SinhalaTokenizer
        tok = SinhalaTokenizer()
        protected = _protect_false_boundaries(text, SINHALA_ABBREVIATIONS)
        sents = tok.split_sentences(protected)
        return [_restore(s).strip() for s in sents if s.strip()]
    except Exception:
        return naive_sentence_split(text, SINHALA_ABBREVIATIONS)


def _tokenize_chunk_tamil(text: str):
    if USE_WTPSPLIT:
        try:
            return _tokenize_chunk_sat(text, TAMIL_ABBREVIATIONS)
        except Exception:
            pass  # fall through to tamil-tokenizer / naive below
    try:
        from tamil_tokenizer import tokenize_sentences
        protected = _protect_false_boundaries(text, TAMIL_ABBREVIATIONS)
        sents = tokenize_sentences(protected)
        return [_restore(s).strip() for s in sents if s.strip()]
    except Exception:
        return naive_sentence_split(text, TAMIL_ABBREVIATIONS)


def _tokenize_raw_sat(text: str):
    """SaT with NO masking, NO pre-chunking - completely raw input, for
    testing how good SaT is on its own."""
    sat = _get_sat_model()
    return [s.strip() for s in sat.split(text) if s.strip()]


def _reattach_marker(marker, sentences):
    """Prepend a numbered-list marker (e.g. '02.') back onto the first
    tokenized sentence of its chunk, verbatim - see split_on_numbered_markers
    for why the tokenizer never sees the marker in the first place."""
    if not marker:
        return sentences
    if sentences:
        sentences[0] = f"{marker} {sentences[0]}"
        return sentences
    return [marker]


def split_sentences_sinhala(text: str):
    if not USE_CUSTOM_RULES:
        if USE_WTPSPLIT:
            try:
                return _tokenize_raw_sat(text)
            except Exception:
                pass
        return naive_sentence_split(text, [])  # no abbreviation list either

    all_sents = []
    for marker, chunk in split_on_numbered_markers(text):
        all_sents.extend(_reattach_marker(marker, _tokenize_chunk_sinhala(chunk)))
    return merge_short_fragments(all_sents)


def split_sentences_tamil(text: str):
    if not USE_CUSTOM_RULES:
        if USE_WTPSPLIT:
            try:
                return _tokenize_raw_sat(text)
            except Exception:
                pass
        return naive_sentence_split(text, [])  # no abbreviation list either

    all_sents = []
    for marker, chunk in split_on_numbered_markers(text):
        all_sents.extend(_reattach_marker(marker, _tokenize_chunk_tamil(chunk)))
    return merge_short_fragments(all_sents)


def build_sentence_list(text: str, lang: str):
    """Returns list of (sid, sentence_text) for the whole document,
    sid = 'P{n}.S{m}'."""
    splitter = split_sentences_sinhala if lang == "si" else split_sentences_tamil
    out = []
    for p_idx, segments in enumerate(split_paragraph_segments(text), start=1):
        s_idx = 1
        for seg in segments:
            for s in splitter(seg):
                out.append((f"P{p_idx}.S{s_idx}", s))
                s_idx += 1
    return out


# ---------------------------------------------------------------------------
# MAIN - build the side-by-side sheet for this one file pair
# ---------------------------------------------------------------------------

def run(sinhala_file: str, tamil_file: str, output_dir: str, output_filename: str):
    os.makedirs(output_dir, exist_ok=True)
    output_xlsx = os.path.join(output_dir, output_filename)

    si_sents = build_sentence_list(load_text(sinhala_file), "si")
    ta_sents = build_sentence_list(load_text(tamil_file), "ta")

    max_len = max(len(si_sents), len(ta_sents))
    rows = []
    for i in range(max_len):
        si_id, si_text = si_sents[i] if i < len(si_sents) else ("", "")
        ta_id, ta_text = ta_sents[i] if i < len(ta_sents) else ("", "")
        rows.append({
            "idx": i + 1,
            "si_id": si_id, "si_sentence": si_text,
            "ta_id": ta_id, "ta_sentence": ta_text,
        })

    df = pd.DataFrame(rows)

    def _write_excel(path):
        df.to_excel(path, index=False, engine="openpyxl")

    import time
    for attempt in range(3):
        try:
            _write_excel(output_xlsx)
            break
        except PermissionError:
            if attempt < 2:
                print(f"  [warning] '{output_xlsx}' appears to be open (e.g. in Excel). "
                      f"Close it now - retrying in 3s...")
                time.sleep(3)
                continue
            fallback_xlsx = output_xlsx.replace(".xlsx", f"_{int(time.time())}.xlsx")
            _write_excel(fallback_xlsx)
            print(f"  [warning] '{output_xlsx}' was locked; wrote to '{fallback_xlsx}' instead.")
            output_xlsx = fallback_xlsx

    # also write a plain-text side-by-side view, and print it to the console,
    # so you can inspect it without opening Excel at all
    txt_path = os.path.splitext(output_xlsx)[0] + ".txt"
    with open(txt_path, "w", encoding="utf-8") as f:
        for i in range(max_len):
            si_id, si_text = si_sents[i] if i < len(si_sents) else ("", "")
            ta_id, ta_text = ta_sents[i] if i < len(ta_sents) else ("", "")
            line = f"[{i + 1}] {si_id:<8} SI: {si_text}\n    {ta_id:<8} TA: {ta_text}\n"
            f.write(line + "\n")

    print("\n" + "=" * 100)
    print(f"{'#':<4}{'SI_ID':<8} SI_SENTENCE")
    print(f"{'':<4}{'TA_ID':<8} TA_SENTENCE")
    print("=" * 100)
    for i in range(max_len):
        si_id, si_text = si_sents[i] if i < len(si_sents) else ("", "")
        ta_id, ta_text = ta_sents[i] if i < len(ta_sents) else ("", "")
        print(f"{i + 1:<4}{si_id:<8} {si_text}")
        print(f"{'':<4}{ta_id:<8} {ta_text}")
        print("-" * 100)

    print(f"Sinhala file: {sinhala_file}")
    print(f"Tamil file  : {tamil_file}")
    print(f"Sinhala sentence count: {len(si_sents)}")
    print(f"Tamil sentence count  : {len(ta_sents)}")
    print(f"Done. Side-by-side sentence split report written to: {output_xlsx}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Sinhala-Tamil sentence-splitting inspector (single file pair)")
    parser.add_argument("--sinhala_file", default=SINHALA_FILE)
    parser.add_argument("--tamil_file", default=TAMIL_FILE)
    parser.add_argument("--output_dir", default=OUTPUT_DIR)
    parser.add_argument("--output_filename", default=OUTPUT_FILENAME)
    args = parser.parse_args()

    run(args.sinhala_file, args.tamil_file, args.output_dir, args.output_filename)