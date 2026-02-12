"""
Fix: Picture/video PA detection and processing

Three bugs fixed:
1. _has_picture_video() was called on `summary` AFTER _normalize_text_preserve()
   stripped the "picture/video was provided" sentence. Now checks `raw_summary`.
2. _format_analysis_block() called _normalize_text_preserve() (which runs boilerplate
   stripping) BEFORE picture processing, so _strip_picture_boilerplate() couldn't
   find the full sentences. Now picture processing runs on raw text first.
3. "Visual inspection:" rename left duplicates. Replaced 4 separate functions with
   one unified _process_picture_pa() that handles everything in correct order.

Usage: python fix_picture_video.py scrape_and_generate.py
"""
import sys, re

def apply_fixes(code: str) -> str:
    fixes = []

    # ================================================================
    # FIX 1: Replace the 4 separate picture functions with unified one
    # ================================================================
    old_funcs = '''def _has_picture_video(text: str) -> bool:
    if not text:
        return False
    return bool(re.search(r'picture/video\\s+was\\s+provided', text, re.I))
def _strip_picture_boilerplate(text: str) -> str:
    if not text:
        return text
    text = re.sub(
        r'The\\s+product\\s+sample\\s+was\\s+not\\s+returned\\s+to\\s+the\\s+Returned\\s+Product\\s+Analysis\\s*\\(RPA\\)\\s*laboratory[;,]?\\s*however[,]?\\s*a\\s+picture/video\\s+was\\s+provided\\s+by\\s+the\\s+customer\\s+for\\s+analysis\\.?\\s*',
        '', text, flags=re.I
    )
    text = re.sub(
        r'A\\s+visual\\s+inspection\\s+of\\s+the\\s+returned\\s+photo\\(s\\)\\s+noted:\\s*',
        'Picture Evaluation:\\n',
        text, flags=re.I
    )
    text = re.sub(r'[ \\t]{2,}', ' ', text)
    text = re.sub(r'\\n{3,}', '\\n\\n', text)
    return text.strip()
def _rename_visual_to_picture_eval(text: str) -> str:
    if not text:
        return text
    return re.sub(
        r'(?m)^(\\s*)Visual\\s+[Ii]nspection\\s*:',
        r'\\1Picture Evaluation:',
        text
    )
def _strip_evaluation_section(text: str) -> str:
    if not text:
        return text
    BULLET = '\\u2022'
    lines = text.split('\\n')
    result = []
    in_eval_section = False
    for line in lines:
        stripped = line.strip()
        if re.match(r'^Evaluation\\s*:$', stripped, re.I):
            in_eval_section = True
            continue
        if in_eval_section:
            if stripped.startswith(BULLET) or stripped.startswith('- ') or stripped.startswith('\\u2022'):
                continue
            if not stripped:
                continue
            in_eval_section = False
            result.append(line)
        else:
            result.append(line)
    return '\\n'.join(result)'''

    new_funcs = '''def _has_picture_video(text: str) -> bool:
    """Detect if a PA summary mentions picture/video was provided."""
    if not text:
        return False
    return bool(re.search(r'picture/video\\s+was\\s+provided', text, re.I))
def _process_picture_pa(text: str) -> str:
    """All picture-specific transformations in one pass.
    Must run BEFORE general boilerplate stripping so it can find
    the full RPA laboratory sentence and photo inspection header."""
    if not text:
        return text
    BULLET = '\\u2022'
    # 1. Remove RPA lab + picture/video sentence
    text = re.sub(
        r'The\\s+product\\s+sample\\s+was\\s+not\\s+returned\\s+to\\s+the\\s+Returned\\s+Product\\s+Analysis\\s*\\(RPA\\)\\s*laboratory[;,]?\\s*however[,]?\\s*a\\s+picture/video\\s+was\\s+provided\\s+by\\s+the\\s+customer\\s+for\\s+analysis\\.?\\s*',
        '', text, flags=re.I)
    # 2. Remove "Visual inspection:" line entirely (replaced by Picture Evaluation)
    text = re.sub(r'(?m)^[ \\t]*Visual\\s+[Ii]nspection\\s*:\\s*\\n?', '', text)
    text = re.sub(r'(?<=[.!?])\\s*Visual\\s+[Ii]nspection\\s*:\\s*', '\\n', text)
    # 3. Replace "A visual inspection of the returned photo(s) noted:" -> "Picture Evaluation:"
    text = re.sub(
        r'A\\s+visual\\s+inspection\\s+of\\s+the\\s+returned\\s+photo\\(s\\)\\s+noted:\\s*',
        'Picture Evaluation:\\n', text, flags=re.I)
    # 4. Remove standalone "Evaluation:" section and its bullet points
    lines = text.split('\\n')
    result = []
    in_eval = False
    for line in lines:
        s = line.strip()
        if re.match(r'^Evaluation\\s*:$', s, re.I):
            in_eval = True
            continue
        if in_eval:
            if s.startswith(BULLET) or s.startswith('- ') or not s:
                continue
            in_eval = False
            result.append(line)
        else:
            result.append(line)
    text = '\\n'.join(result)
    text = re.sub(r'[ \\t]{2,}', ' ', text)
    text = re.sub(r'\\n{3,}', '\\n\\n', text)
    return text.strip()'''

    if old_funcs in code:
        code = code.replace(old_funcs, new_funcs, 1)
        fixes.append("Replaced 4 picture functions with _has_picture_video + _process_picture_pa")
    else:
        print("  [WARN] Could not find old picture functions block")

    # ================================================================
    # FIX 2: Update _format_analysis_block to run picture processing
    #         on raw text BEFORE _normalize_text_preserve
    # ================================================================
    old_format = '''def _format_analysis_block(product_desc: str, summary: str, product_count: int = 1, include_lead: bool = True, has_picture: bool = False) -> str:
    s = "" if summary is None else _normalize_text_preserve(summary)
    s = _replace_dashes_with_bullets(s)
    if has_picture:
        s = _strip_picture_boilerplate(s)
        s = _strip_evaluation_section(s)
        s = _rename_visual_to_picture_eval(s)'''

    new_format = '''def _format_analysis_block(product_desc: str, summary: str, product_count: int = 1, include_lead: bool = True, has_picture: bool = False) -> str:
    s = summary or ""
    if has_picture:
        s = _process_picture_pa(s)
    s = _normalize_text_preserve(s)
    s = _replace_dashes_with_bullets(s)'''

    if old_format in code:
        code = code.replace(old_format, new_format, 1)
        fixes.append("Fixed _format_analysis_block: picture processing now runs BEFORE boilerplate")
    else:
        print("  [WARN] Could not find _format_analysis_block to fix")

    # ================================================================
    # FIX 3: Check _has_picture_video on raw_summary, not summary
    # ================================================================
    old_check = '''            is_picture = _has_picture_video(summary)
            log(f"[PA-SUMMARY] txid={txid} is_picture={is_picture}")'''

    new_check = '''            is_picture = _has_picture_video(raw_summary)
            log(f"[PA-SUMMARY] txid={txid} is_picture={is_picture}")'''

    if old_check in code:
        code = code.replace(old_check, new_check, 1)
        fixes.append("Fixed picture detection: now checks raw_summary before boilerplate stripping")
    else:
        print("  [WARN] Could not find is_picture check to fix")

    print(f"\n  Applied {len(fixes)} fixes:")
    for f in fixes:
        print(f"    [OK] {f}")

    return code


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python fix_picture_video.py <path_to_scrape_and_generate.py>")
        sys.exit(1)

    inpath = sys.argv[1]
    with open(inpath, 'r', encoding='utf-8') as f:
        code = f.read()

    print(f"Fixing {inpath}...")
    fixed = apply_fixes(code)

    # Write back to same file
    with open(inpath, 'w', encoding='utf-8') as f:
        f.write(fixed)

    print(f"\nFixes applied in-place to: {inpath}")