import re, sys, time, json
from pathlib import Path
from datetime import date
import yaml
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from docx import Document
def ts():
    return time.strftime("%Y-%m-%d %H:%M:%S")
def log(msg):
    print(f"[{ts()}] {msg}")
def find_first_visible_input(page, primary_selector, fallbacks=None, timeout=15000):
    fallbacks = fallbacks or []
    selectors = [primary_selector] + fallbacks
    contexts = [page] + [fr for fr in page.frames]
    last_exc = None
    for sel in selectors:
        for ctx in contexts:
            try:
                loc = ctx.locator(sel).first
                try:
                    loc.wait_for(state="visible", timeout=timeout)
                    return loc, ctx
                except Exception as e:
                    last_exc = e
                count = loc.count()
                if count:
                    return loc, ctx
            except Exception as e:
                last_exc = e
                continue
    raise PWTimeout(f"Could not find element (visible or attached) for selectors: {selectors}")
def dump_frames_debug(page, basename="debug"):
    info_lines = []
    frames = page.frames
    for idx, fr in enumerate(frames):
        try:
            url = fr.url or ""
            name = fr.name or ""
        except Exception:
            url = ""
            name = ""
        info_lines.append(f"[{idx}] name={name!r} url={url!r}")
        try:
            pass
        except Exception:
            pass
        try:
            candidates = fr.locator("xpath=//input[contains(@class,'th-sif') or contains(@id,'SearchValue') or contains(@tempname,'search')]")
            count = candidates.count()
            info_lines.append(f"  candidates(th-sif/id*SearchValue/tempname*search): {count}")
            for j in range(min(count, 10)):
                el = candidates.nth(j)
                html = el.evaluate("e => e.outerHTML").strip()
                info_lines.append("    " + html.replace("\n"," "))
        except Exception as e:
            info_lines.append(f"  (error collecting candidates: {e})")
    try:
        Path(f"{basename}_frames.txt").write_text("\n".join(info_lines), encoding="utf-8")
    except Exception:
        pass
def soft_click_go(ctx, custom_selector=None):
    candidates = []
    if custom_selector:
        candidates.append(custom_selector)
    candidates += [
        "text=Go",
        "input[value='Go']",
        "button:has-text('Go')",
        "xpath=//a[normalize-space(.)='Go']",
    ]
    for sel in candidates:
        try:
            btn = ctx.locator(sel).first
            if btn.count():
                btn.click()
                return True
        except Exception:
            pass
    return False
def wait_find_in_any_frame(page, selectors, timeout_ms=30000, poll_ms=600):
    import time
    deadline = time.time() + (timeout_ms/1000.0)
    tried = set()
    while time.time() < deadline:
        log(f"[wait] scanning {len(page.frames)} frames for any of: {selectors}")
        frames = page.frames
        for sel in selectors:
            if sel in tried:
                pass
            for fr in frames:
                try:
                    loc = fr.locator(sel).first
                    if loc.count():
                        try:
                            loc.wait_for(state="visible", timeout=poll_ms)
                        except Exception:
                            pass
                        log(f"[wait] found {sel} in frame name={getattr(fr,'name','')} url={getattr(fr,'url','')}")
                        return loc, fr, sel
                except Exception:
                    continue
        time.sleep(poll_ms/1000.0)
    raise PWTimeout(f"Could not find element in any frame for selectors: {selectors}")
import re
from docx import Document
def _norm_key(s: str) -> str:
    return re.sub(r'[^a-z0-9]+', '_', (s or '').strip().lower()).strip('_')
_PLACEHOLDER_ANY = re.compile(r'(\{\{|\[\[)\s*(.*?)\s*(\}\}|\]\])', re.I)
def _build_alias_mapping(mapping: dict) -> dict:
    out = {}
    for k, v in mapping.items():
        nk = _norm_key(k)
        out[nk] = v
        if nk.endswith('_1'):
            out[nk[:-2]] = v  
    aliases = {
        'today_date': out.get('todays_date', ''),
        'ir_name': out.get('ir_name', ''),
        'ir_with_address': out.get('ir_with_address', ''),
        'event_date': out.get('event_date', ''),
        'event_description': out.get('event_description', ''),
        'analysis_results_if_present': out.get('analysis_results', ''),
        'investigation_summary': out.get('investigation_summary', ''),
        'product_id': out.get('product_id_1', ''),
        'product_desc': out.get('product_desc_1', ''),
        'lot_serial_number': out.get('serial_or_lot_1', ''),
        'lot_serial_no': out.get('serial_or_lot_1', ''),
        'serial_no_lot_no': out.get('serial_or_lot_1', ''),
    }
    out.update({k: v for k, v in aliases.items() if v})
    return out
def _split_tolerant(label: str) -> str:
    label = re.sub(r'\s+', ' ', label.strip())
    parts = []
    gap = r'(?:\s*<\/w:t>\s*<\/w:r>\s*<w:r[^>]*>\s*<w:t[^>]*>\s*)?'
    for ch in label:
        if ch == ' ':
            parts.append(r'\s+')
        else:
            parts.append(re.escape(ch))
        parts.append(gap)
    return ''.join(parts)
def _patterns_for_key(human_label: str):
    inner = _split_tolerant(human_label)
    p1 = re.compile(rf'\[\[\s*{inner}\s*\]\]', re.I)
    p2 = re.compile(rf'\[\[\s*{inner}\s*\]\]', re.I)
    return p1, p2
def _xml_replace_all(xml: str, mapping: dict) -> str:
    def _quick(m):
        return mapping.get(_norm_key(m.group(2)), '')
    xml = _PLACEHOLDER_ANY.sub(lambda m: _quick(m), xml)
    keys_seen = set()
    for raw_key, value in mapping.items():
        if not value:
            continue
        humanish = raw_key.replace('_', ' ').strip()
        for label in {raw_key, humanish}:
            if label in keys_seen: 
                continue
            keys_seen.add(label)
            for pat in _patterns_for_key(label):
                xml = pat.sub(value, xml)
    return xml
def replace_everywhere(doc: Document, mapping: dict):
    resolved = _build_alias_mapping(mapping)
    pkg = doc.part.package
    for part in pkg.parts:   # includes document.xml, headers, footers, footnotes, etc.
        ct = getattr(part, 'content_type', '')
        if not ct or 'xml' not in ct:
            continue
        try:
            xml = part.blob.decode('utf-8')
        except Exception:
            continue
        new_xml = _xml_replace_all(xml, resolved)
        if new_xml != xml:
            part._blob = new_xml.encode('utf-8')
def fill_docx(template_path, out_path, mapping):
    doc = Document(template_path)
    replace_everywhere(doc, mapping)
    doc.save(out_path)
def find_app_frame(page, frame_name_regex=None, url_regex=None):
    if url_regex:
        for fr in page.frames:
            if re.search(url_regex, fr.url or "", re.I):
                return fr
    if frame_name_regex:
        for fr in page.frames:
            if re.search(frame_name_regex, (fr.name or ""), re.I):
                return fr
    return page.main_frame
def clean(s):
    return re.sub(r'\s{2,}', ' ', (s or '').replace('\xa0', ' ')).strip()
def get_by_label(frame, labels):
    if isinstance(labels, str):
        labels = [labels]
    for label in labels:
        loc = frame.locator(f"xpath=//*[normalize-space(text())='{label}']")
        if loc.count() > 0:
            try:
                el = loc.first
                sib = el.locator("xpath=following-sibling::*[1]")
                if sib.count():
                    return clean(sib.first.inner_text())
            except Exception:
                pass
            try:
                for_attr = loc.first.get_attribute('for')
                if for_attr:
                    inp = frame.locator(f"#{for_attr}")
                    if inp.count():
                        val = (inp.first.input_value()
                            if 'input' in (inp.first.evaluate('e => e.tagName') or '').lower()
                            else inp.first.inner_text())
                        return clean(val)
            except Exception:
                pass
            try:
                nxt = loc.first.locator("xpath=following::*[1]")
                if nxt.count():
                    return clean(nxt.first.inner_text())
            except Exception:
                pass
        loc2 = frame.locator(f"xpath=//*[contains(normalize-space(.), '{label}')]")
        if loc2.count():
            try:
                sib = loc2.first.locator("xpath=following-sibling::*[1]")
                if sib.count():
                    return clean(sib.first.inner_text())
            except Exception:
                pass
    return ""
def get_grid_row_col(frame, grid_hint_xpath, row_index, col_index):
    try:
        grid = frame.locator(grid_hint_xpath).first
        if grid.count() == 0:
            return ""
        rows = grid.locator("xpath=.//tr")
        row = rows.nth(row_index)
        cells = row.locator("xpath=.//th|.//td")
        cell = cells.nth(col_index - 1)
        return clean(cell.inner_text())
    except Exception:
        return ""
def click_partners_tab(page, frame):
    sels = [
        "xpath=//a[contains(@class,'ui-tabs-anchor')][normalize-space(.)='Partners']",
        "text=Partners",
        "xpath=//a[@class='ui-tabs-anchor' and contains(@href,'_ovviewset.do_0008')]",
    ]
    clicked = False
    for sel in sels:
        try:
            loc = frame.locator(sel).first
            if loc.count():
                log("[nav] Clicking Partners tab")
                loc.click()
                clicked = True
                break
        except Exception:
            pass
    if not clicked:
        try:
            loc, ctx, _ = wait_find_in_any_frame(page, sels, timeout_ms=15000)
            loc.click()
            clicked = True
        except Exception:
            pass
    if not clicked:
        return False  # nothing to click
    grid_sigs = [
        "xpath=//td[starts-with(@id,'GUIDE-PartnersTable-')]",
        "xpath=//td[@aria-label='Partner Function']",
        "xpath=//th[normalize-space(.)='Partner Function']",
    ]
    try:
        wait_find_in_any_frame(page, grid_sigs, timeout_ms=20000)
    except Exception:
        pass
    return True
def get_initial_reporter_name(frame):
    tr = _row_by_pf_in_partners(frame, ["Initial Reporter", "Initial Contact", "Initial Reporter/Contact"])
    if not tr:
        return ""
    return _cell_text_in_same_row(tr, "Name")
def get_facility_name_and_address(frame):
    tr = _row_by_pf_in_partners(frame, ["Facility", "Health Care Facility", "Healthcare Facility", "Plant"])
    if not tr:
        return ""
    name = _cell_text_in_same_row(tr, "Name")
    addr = _cell_text_in_same_row(tr, "Address") or _cell_text_in_same_row(tr, "address_short")
    return f"{name}\n{addr}".strip()
def _partners_table(frame):
    t = frame.locator("xpath=//table[.//td[starts-with(@id,'GUIDE-PartnersTable-')]]").first
    return t if t.count() else None
def _row_by_pf_in_partners(frame, pf_names):
    tbl = _partners_table(frame)
    if not tbl:
        return None
    if isinstance(pf_names, str):
        pf_names = [pf_names]
    for pf in pf_names:
        tr = tbl.locator(
            "xpath=.//tr[td[starts-with(@id,'GUIDE-PartnersTable-') and "
            "contains(@id,'-PartnerFunction') and normalize-space(.)=$pf]]"
        ).filter(has_text=pf).first
        if tr.count():
            return tr
        tr = tbl.locator(
            f"xpath=.//tr[td[starts-with(@id,'GUIDE-PartnersTable-') and contains(@id,'-PartnerFunction') "
            f"and normalize-space(.)='{pf}']]"
        ).first
        if tr.count():
            return tr
    for pf in pf_names:
        low = pf.lower()
        tr = tbl.locator(
            "xpath=.//tr[td[starts-with(@id,'GUIDE-PartnersTable-') and contains(@id,'-PartnerFunction') "
            f"and contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{low}')]]"
        ).first
        if tr.count():
            return tr
    return None
def find_partners_frame(page, timeout_ms=10000, poll_ms=200):
    import time
    deadline = time.time() + (timeout_ms / 1000.0)
    sel = ("xpath=//td[starts-with(@id,'GUIDE-PartnersTable-') or "
           "starts-with(@id,'C21_W80_V81_btpartner_table')]")
    while time.time() < deadline:
        for fr in page.frames:
            try:
                loc = fr.locator(sel).first
                if loc.count():
                    return fr
            except Exception:
                pass
        time.sleep(poll_ms / 1000.0)
    return None
def debug_frames_for_partners(page):
    print("[Partners] Scanning frames for partners grid…")
    for i, fr in enumerate(page.frames):
        try:
            has = fr.locator("xpath=//td[starts-with(@id,'GUIDE-PartnersTable-') or starts-with(@id,'C21_W80_V81_btpartner_table')]").count()
            print(f"  [{i}] name={fr.name!r} url={fr.url!r}  matches={has}")
        except Exception as e:
            print(f"  [{i}] error: {e}")
def _cell_text_in_same_row(tr_loc, col_name):
    td = tr_loc.locator(
        f"xpath=.//td[starts-with(@id,'GUIDE-PartnersTable-') and contains(@id,'-{col_name}')]"
    ).first
    if not td.count():
        return ""
    return clean(td.inner_text())
def _debug_list_pf_from_correct_table(frame):
    tbl = _partners_table(frame)
    if not tbl:
        print("[Partners] Could not find the Partners table (GUIDE-PartnersTable).")
        return
    cells = tbl.locator("xpath=.//td[starts-with(@id,'GUIDE-PartnersTable-') and contains(@id,'-PartnerFunction')]")
    n = cells.count()
    print(f"[Partners] Partner table rows detected: {n}")
    for i in range(min(n, 30)):
        print("  -", clean(cells.nth(i).inner_text()))
def find_frame_with(page, selector, timeout_ms=10000, poll_ms=200):
    import time
    deadline = time.time() + (timeout_ms/1000.0)
    while time.time() < deadline:
        for fr in page.frames:
            try:
                if fr.locator(selector).first.count():
                    return fr
            except Exception:
                pass
        time.sleep(poll_ms/1000.0)
    return None
def click_tab_by_text(page, root_frame, text_or_href_fragment):
    sels = [
        f"xpath=//a[contains(@class,'ui-tabs-anchor') and normalize-space(.)='{text_or_href_fragment}']",
        f"text={text_or_href_fragment}",
        f"xpath=//a[contains(@class,'ui-tabs-anchor') and contains(@href,'{text_or_href_fragment}')]",
    ]
    for fr in [root_frame] + list(page.frames):
        for sel in sels:
            try:
                loc = fr.locator(sel).first
                if loc.count():
                    log(f"[nav] Clicking tab: {text_or_href_fragment}")
                    loc.click()
                    return fr
            except Exception:
                pass
    return None
_COMPLETE_RX = re.compile(
    r"\b(complete(?:d)?|closed\s*[-–]?\s*complete|closure\s*[-–]?\s*complete|final(?:ized)?|fully\s*resolved)\b",
    re.I,
)
_ID_PATTERNS = [
    r"\b\d{5,}\b",                 # plain long numbers
    r"\bWI[-_ ]?\d{4,}\b",
    r"\bINV[-_ ]?\d{4,}\b",
    r"\bPA[-_ ]?\d{4,}\b",
    r"\bTXN[-_ ]?\d{4,}\b",
    r"\bAN[-_ ]?\d{4,}\b",
    r"\b[A-Z]{2,5}[-_ ]?\d{4,}\b", # generic code-12345
]
def _find_first_match(patterns, s):
    for p in patterns:
        m = re.search(p, s or "", re.I)
        if m:
            return m.group(0)
    return None
def _row_combined_text(row):
    try:
        return clean(row.inner_text())
    except Exception:
        try:
            return (row.evaluate("n => n.textContent || ''") or "").strip()
        except Exception:
            return ""
def _row_txid(row):
    combined = _row_combined_text(row)
    txid = _find_first_match(_ID_PATTERNS, combined)
    if txid:
        return txid
    links = row.locator("xpath=.//a[@href]")
    for k in range(links.count()):
        href = (links.nth(k).get_attribute("href") or "")
        txid = _find_first_match(_ID_PATTERNS, href)
        if txid:
            return txid
        m = re.search(r"(?:id|no|number|case|txn|wi)[=:/#](\w[-\w]*)", href, re.I)
        if m:
            return m.group(1)
    cells = row.locator("xpath=.//th|.//td")
    for j in range(cells.count()):
        el = cells.nth(j)
        for attr in ("data-id", "data-transactionid", "data-txid", "data-key"):
            val = (el.get_attribute(attr) or "")
            txid = _find_first_match(_ID_PATTERNS, val)
            if txid:
                return txid
    return None
def _row_status_text(row):
    status_like = row.locator(
        "xpath=.//*[(self::td or self::th) and "
        " (contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'status') or "
        "  contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'state') or "
        "  contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'result') or "
        "  contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'outcome') or "
        "  contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'disposition') or "
        "  contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'resolution'))]"
    ).first
    if status_like.count():
        t = clean(status_like.inner_text()) or ""
        if t:
            return t
    hinted = row.locator("xpath=.//*[@title or @aria-label or @alt]")
    for k in range(hinted.count()):
        node = hinted.nth(k)
        for attr in ("title", "aria-label", "alt"):
            val = node.get_attribute(attr) or ""
            if _COMPLETE_RX.search(val):
                return val
    return _row_combined_text(row)
def _row_is_complete(row):
    t = (_row_status_text(row) or "").lower()
    if _COMPLETE_RX.search(t):
        return True
    return ("closed" in t and "complete" in t)
def _pli_table(frame):
    return frame.locator("xpath=//table[.//td[starts-with(@id,'GUIDE-ProductLineItemsTable-')]]").first
def _get_attr_or_text(node):
    return clean(
        (node.get_attribute("title") or node.get_attribute("aria-label") or node.inner_text() or "")
    )
def read_all_products(page, root_frame):
    click_tab_by_text(page, root_frame, "Product Line Items") or \
    click_tab_by_text(page, root_frame, "_ovviewset.do_0002")
    fr = find_frame_with(page, "xpath=//td[starts-with(@id,'GUIDE-ProductLineItemsTable-')]")
    if fr:
        tbl = _pli_table(fr)
        if tbl and tbl.count():
            rows = tbl.locator(
                "xpath=.//tr[td[starts-with(@id,'GUIDE-ProductLineItemsTable-') and contains(@id,'-Product')]]"
            )
            n = rows.count()
            out = []
            for i in range(n):
                row = rows.nth(i)
                prod = row.locator(
                    "xpath=.//td[starts-with(@id,'GUIDE-ProductLineItemsTable-') and contains(@id,'-Product')]"
                ).first
                desc = row.locator(
                    "xpath=.//td[starts-with(@id,'GUIDE-ProductLineItemsTable-') and contains(@id,'-Description')]"
                ).first
                pid = clean(prod.inner_text()) if prod.count() else ""
                pdesc = clean(desc.inner_text()) if desc.count() else ""
                pcode = extract_product_code(pdesc)
                sn_candidates = [
                    "xpath=.//td[starts-with(@id,'GUIDE-ProductLineItemsTable-') and contains(@id,'-SN')]",
                    "xpath=.//span[contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'s/n')]",
                    "xpath=.//span[contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'serial')]",
                ]
                sn_val = ""
                for sel in sn_candidates:
                    sn_el = row.locator(sel).first
                    if sn_el.count():
                        sn_val = clean(sn_el.inner_text())
                        if sn_val:
                            break
                lot_candidates = [
                    "xpath=.//td[starts-with(@id,'GUIDE-ProductLineItemsTable-') and contains(@id,'-Lot')]",
                    "xpath=.//span[contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'lot')]",
                ]
                lot_val = ""
                for sel in lot_candidates:
                    lot_el = row.locator(sel).first
                    if lot_el.count():
                        lot_val = clean(lot_el.inner_text())
                        if lot_val:
                            break
                if pid or pdesc or sn_val or lot_val:
                    out.append({"id": pid, "desc": pdesc, "code": pcode, "sn": sn_val, "lot": lot_val})
            if out:
                return out  # Successfully parsed GUIDE table; we're done.
    fr = find_frame_with(page, "xpath=//*[contains(@id,'btadmini_table')]")
    if not fr:
        return []
    rows = fr.locator(
        "xpath=.//tr[.//a[contains(@id,'ordered_prod')] or .//a[contains(@id,'number_int')]]"
    )
    n = rows.count()
    out = []
    for i in range(n):
        row = rows.nth(i)
        ordered_link = row.locator("xpath=.//a[contains(@id,'ordered_prod')]").first
        pid = ""
        if ordered_link.count():
            try:
                pid = (
                    ordered_link.get_attribute("title")
                    or ordered_link.get_attribute("aria-label")
                    or ordered_link.inner_text()
                    or ""
                )
                pid = clean(pid)
            except Exception:
                pid = ""
        pdesc = ""
        try:
            if ordered_link.count():
                maybe_desc = ordered_link.locator("xpath=ancestor::td[1]/following-sibling::td[1]").first
                if maybe_desc.count():
                    pdesc = clean(maybe_desc.inner_text())
        except Exception:
            pass
        pcode = extract_product_code(pdesc) or pid
        sn_val = ""
        lot_val = ""
        if pid or pdesc:
            out.append({"id": pid, "desc": pdesc, "code": pcode, "sn": sn_val, "lot": lot_val})
    return out
def _dates_table(frame):
    return frame.locator("xpath=//table[.//td[starts-with(@id,'GUIDE-DatesTable')]]").first
def get_event_date(page):
    click_tab_by_text(page, page.main_frame, "Dates") or \
    click_tab_by_text(page, page.main_frame, "_ovviewset.do_0003")
    fr = find_frame_with(page, "xpath=//td[starts-with(@id,'GUIDE-DatesTable')]")
    if not fr:
        return ""
    tbl = _dates_table(fr)
    if not tbl or not tbl.count():
        return ""
    row = tbl.locator(
        "xpath=.//tr[td[starts-with(@id,'GUIDE-DatesTable') and contains(@id,'-DateType') and normalize-space(.)='Event Date']]"
    ).first
    if not row.count():
        return ""
    cell = row.locator(
        "xpath=.//td[starts-with(@id,'GUIDE-DatesTable') and contains(@id,'-DateFrom')]"
    ).first
    return clean(cell.inner_text()) if cell.count() else ""
def _aer_table(frame):
    return frame.locator("xpath=//table[.//td[starts-with(@id,'GUIDE-AdditionalExternalReferencesTable-')]]").first
def _aer_row_by_type(tbl, *type_fragments):
    for t in type_fragments:
        row = tbl.locator(
            f"xpath=.//tr[td[starts-with(@id,'GUIDE-AdditionalExternalReferencesTable-') and contains(@id,'-ExtReferenceType') and normalize-space(.)='{t}']]"
        ).first
        if row.count():
            return row
    for t in type_fragments:
        low = t.lower()
        row = tbl.locator(
            "xpath=.//tr[td[starts-with(@id,'GUIDE-AdditionalExternalReferencesTable-') and contains(@id,'-ExtReferenceType') "
            f"and contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{low}')]]"
        ).first
        if row.count():
            return row
    return None
def _aer_number_from_row(row):
    for suffix in ("-ExtReferenceNumber", "-ExtReferenceID", "reference_number", "reference_id"):
        td = row.locator(
            f"xpath=.//td[starts-with(@id,'GUIDE-AdditionalExternalReferencesTable-') and contains(@id,'{suffix}')]"
        ).first
        if td.count():
            return clean(td.inner_text())
    return ""
def read_external_refs(page, root_frame):
    owner = click_tab_by_text(page, root_frame, "Additional External References") or \
            click_tab_by_text(page, root_frame, "_ovviewset.do_0013")
    fr = find_frame_with(page, "xpath=//td[starts-with(@id,'GUIDE-AdditionalExternalReferencesTable-')]")
    if not fr:
        return {"rb_reference": "", "report_number": ""}
    tbl = _aer_table(fr)
    if not tbl or not tbl.count():
        return {"rb_reference": "", "report_number": ""}
    out = {"rb_reference": "", "report_number": ""}
    rb_row = _aer_row_by_type(tbl, "RB Reference", "RB", "RBReference")
    if rb_row:
        out["rb_reference"] = _aer_number_from_row(rb_row)
    rep_row = _aer_row_by_type(tbl, "MPXR", "MPRR", "Report", "Report Number")
    if rep_row:
        out["report_number"] = _aer_number_from_row(rep_row)
    return out
def _find_latest_text_table(page):
    tables = []
    for fr in page.frames:
        try:
            scope = fr.locator("xpath=//div[contains(@class,'th-clr-cnt-bottom')]")
            scope = scope if scope.count() else fr
            tds = scope.locator("xpath=.//table[.//td[starts-with(@id,'GUIDE-TextInfoTable-')]]")
            n = tds.count()
            for i in range(n):
                tables.append((fr, tds.nth(i)))
        except Exception:
            pass
    return tables[-1] if tables else (None, None)
def read_event_description(page, root_frame):
    click_tab_by_text(page, root_frame, "Text Info") or \
    click_tab_by_text(page, root_frame, "_ovviewset.do_0006")
    frames = active_content_frames(page, {"content_frame_name_regex": "WorkAreaFrame1"})
    tbl = None
    chosen_frame = None
    for fr in frames:
        try:
            fr.wait_for_selector("xpath=//td[starts-with(@id,'GUIDE-TextInfoTable-')]", timeout=3000)
            candidate = _find_latest_text_table_in(fr)
            if candidate:
                tbl, chosen_frame = candidate, fr
                break
        except Exception:
            continue
    if not (tbl and chosen_frame):
        fr, tbl_fallback = _find_latest_text_table(page)
        if not (fr and tbl_fallback and tbl_fallback.count()):
            return ""
        chosen_frame, tbl = fr, tbl_fallback
    want_types = [
        "Incident description", "Incident description / Reason for report",
        "Reason for report", "Description of Event", "Event Description",
        "Narrative", "HCP Narrative", "Event narrative", "Incident Narrative"
    ]
    row = None
    for t in want_types:
        cand = tbl.locator(
            "xpath=.//tr[td[starts-with(@id,'GUIDE-TextInfoTable-') and contains(@id,'-TextType') and normalize-space(.)=$t]]"
        ).filter(has_text=t).first
        if cand.count():
            row = cand
            break
    if not (row and row.count()):
        for t in want_types:
            low = t.lower()
            cand = tbl.locator(
                "xpath=.//tr[td[starts-with(@id,'GUIDE-TextInfoTable-') and contains(@id,'-TextType') "
                f"and contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{low}')]]"
            ).first
            if cand.count():
                row = cand
                break
    if not (row and row.count()):
        row = tbl.locator("xpath=.//tr[td[starts-with(@id,'GUIDE-TextInfoTable-') and contains(@id,'-Text')]]").first
        if not row.count():
            log("[Text] No row with -Text found")
            return ""
    text_td = row.locator(
        "xpath=.//td[starts-with(@id,'GUIDE-TextInfoTable-') and contains(@id,'-Text') and not(contains(@id,'-TextType'))]"
    ).first
    if not text_td.count():
        log("[Text] No -Text cell present in row")
        return ""
    _suppress_clicks_enable(chosen_frame)
    try:
        desc = _safe_td_text(text_td)
    finally:
        _suppress_clicks_disable(chosen_frame)

    log(f"[Text] Event description read from TD (no click), length={len(desc)}")
    return desc
def click_left_nav_product_analysis(page):
    for fr in page.frames:
        try:
            header = fr.locator("xpath=//*[contains(@class,'left-nav')]//*[contains(@class,'ProductAnalysis')]").first
            if not header.count():
                continue
            log("[nav] Opening left nav: Product Analysis")
            _pa_try_expand(fr)  # <-- use the helper
            try:
                fr.wait_for_selector(_section_anchor_xpath("Product Analysis"), timeout=3000)
            except Exception:
                pass
            return True
        except Exception:
            continue
    return False
def click_left_nav_investigation(page):
    candidates = [
        "xpath=//div[contains(@class,'left-nav')]//div[normalize-space(.)='Investigation']",
        "xpath=//*[contains(@class,'left-nav')]//*[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'investigation')]",
    ]
    for sel in candidates:
        for fr in page.frames:
            try:
                loc = fr.locator(sel).first
                if loc.count():
                    log("[nav] Opening left nav: Investigation")
                    loc.click()
                    try:
                        fr.wait_for_selector("xpath=//a[contains(@class,'GUIDE-sideNav')]", timeout=4000)
                    except Exception:
                        pass
                    return True
            except Exception:
                continue
    return False
def _textinfo_signature(page):
    fr, tbl = _find_latest_analysis_table_nearby(page)
    if not (fr and tbl and tbl.count()):
        return ""
    try:
        s = tbl.evaluate("t => (t.innerText || '').slice(0, 800)")
    except Exception:
        s = tbl.inner_text()[:800] if tbl.count() else ""
    return s
def wait_for_textinfo_change(page, previous_sig, timeout=10000):
    start = time.time()
    while time.time() - start < (timeout/1000.0):
        sig = _textinfo_signature(page)
        if sig and sig != previous_sig:
            return True
        time.sleep(0.15)
    return False
def _find_leftnav_frame(page):
    for fr in page.frames:
        try:
            if fr.locator("xpath=//*[contains(@class,'left-nav')]").first.count():
                return fr
        except Exception:
            pass
    for fr in page.frames:
        try:
            if fr.locator("css=a[data-trans-id], a[data-transId], a[data-transid], a.GUIDE-sideNav, a.GUIDE-sideNavLink").first.count():
                return fr
        except Exception:
            pass
    return None
def _section_class(section_text: str) -> str:
    return {
        "Product Analysis": "ProductAnalysis",
        "Investigations":   "Investigation",
        "Product Line Items":"PLI",
    }.get(section_text, re.sub(r"\W+", "", section_text))
def _section_anchor_xpath(section: str) -> str:
    cls = _section_class(section)
    return (
        "xpath=//*[contains(@class,'left-nav')]"
        f"//div[contains(@class,'{cls}')]"
        "/following-sibling::*[contains(@class,'clicker')][1]"
        "/following-sibling::*[contains(@class,'data-wrapper')][1]"
        "//a[(contains(@class,'GUIDE-sideNav') or contains(@class,'GUIDE-sideNavLink') "
        "     or @data-trans-id or @data-transId or @data-transid)]"
    )
def _leftnav_section_header(fr, section_text: str):
    cls = _section_class(section_text)
    header = fr.locator(
        f"xpath=//*[contains(@class,'left-nav')]//div[contains(@class,'{cls}')]"
    ).first
    return header if header.count() else None
def _leftnav_section_container(fr, section_text: str):
    header = _leftnav_section_header(fr, section_text)
    if not header:
        return None
    clicker = header.locator("xpath=following-sibling::div[contains(@class,'clicker')][1]").first
    container = header.locator("xpath=following-sibling::div[contains(@class,'data-wrapper')][1]").first
    if container.count() == 0 and clicker.count():
        robust_click(clicker, fr)
        try:
            fr.wait_for_timeout(250)
        except Exception:
            pass
        container = header.locator("xpath=following-sibling::div[contains(@class,'data-wrapper')][1]").first
    return container if container.count() else None
def _anchors_in_section(fr, section_text: str):
    anchors = fr.locator(_section_anchor_xpath(section_text))
    if anchors.count():
        return anchors
    container = _leftnav_section_container(fr, section_text)
    return container.locator("xpath=.//a[(contains(@class,'GUIDE-sideNav') or contains(@class,'GUIDE-sideNavLink') "
                             "           or @data-trans-id or @data-transId or @data-transid)]") if container else fr.locator("xpath=//*[(0=1)]")
def _enumerate_section_items(fr, section_text: str):
    anchors = _anchors_in_section(fr, section_text)
    n = anchors.count()
    items = []
    for i in range(n):
        a = anchors.nth(i)
        try:
            txt = (a.inner_text().strip() or (a.get_attribute("title") or "").strip())
        except Exception:
            txt = (a.get_attribute("title") or "").strip()
        did = a.get_attribute("data-trans-id") or a.get_attribute("data-transId") or a.get_attribute("data-transid") or ""
        code = extract_product_code(txt)
        items.append({"i": i, "text": txt, "code": code, "data_id": (did or "").strip()})
    log(f"[LeftNav:{section_text}] anchors detected: {n}")
    return items
def _scan_pa_anywhere(page, section_text: str):
    found = []
    for fr in page.frames:
        try:
            anchors = fr.locator(
                "xpath=//*[contains(@class,'left-nav')]"
                "//div[contains(@class,'ProductAnalysis')]"
                "/following-sibling::*[contains(@class,'clicker')][1]"
                "/following-sibling::*[contains(@class,'data-wrapper')][1]"
                "//a[(contains(@class,'GUIDE-sideNav') or contains(@class,'GUIDE-sideNavLink') "
                "     or @data-trans-id or @data-transId or @data-transid)]"
            )
            n = anchors.count()
            if n == 0:
                anchors = fr.locator(
                    "xpath=//*[contains(@class,'left-nav')]"
                    "//a[(contains(@class,'GUIDE-sideNav') or contains(@class,'GUIDE-sideNavLink') "
                    "     or @data-trans-id or @data-transId or @data-transid)]"
                )
                n = anchors.count()
            for i in range(n):
                a = anchors.nth(i)
                try:
                    txt = (a.inner_text().strip() or (a.get_attribute('title') or '')).strip()
                except Exception:
                    txt = (a.get_attribute('title') or '').strip()
                did = (a.get_attribute('data-trans-id') or a.get_attribute('data-transId') or a.get_attribute('data-transid') or '').strip()
                found.append((fr, a, txt, did))
        except Exception:
            continue
    return found
def _product_analysis_anchor_locator(fr):
    return fr.locator(_section_anchor_xpath("Product Analysis"))
def _enumerate_pa_items(fr):
    return _enumerate_section_items(fr, "Product Analysis")
def _content_frame(page):
    for fr in page.frames:
        if (fr.name or "") == "WorkAreaFrame1":
            return fr
    return page.main_frame
def _ensure_section_expanded(page, section: str):
    fr = _find_leftnav_frame(page)
    if not fr:
        return
    if section == "Product Analysis":
        _pa_try_expand(fr)
        try:
            fr.wait_for_selector(_section_anchor_xpath("Product Analysis"), timeout=1200)
        except Exception:
            pass
        return
    cls = _section_class(section)
    header = fr.locator(
        f"xpath=//*[contains(@class,'left-nav')]//*[contains(@class,'{cls}')]"
    ).first
    if not header.count():
        return
    clicker = header.locator("xpath=following-sibling::*[contains(@class,'clicker')][1]").first
    container = header.locator("xpath=following-sibling::*[contains(@class,'data-wrapper')][1]").first
    need_click = True
    if container.count():
        try:
            need_click = container.evaluate("n => getComputedStyle(n).display === 'none'")
        except Exception:
            pass
    if need_click and clicker.count():
        robust_click(clicker, fr)
        try: fr.wait_for_timeout(250)
        except Exception: pass
def list_side_nav_items(page, section: str):
    fr = _find_leftnav_frame(page)
    if not fr:
        return []
    _ensure_section_expanded(page, section)
    anchors = fr.locator(_section_anchor_xpath(section))
    n = anchors.count()
    out = []
    for i in range(n):
        el = anchors.nth(i)
        try:
            t = clean(el.inner_text())
            if not t:
                t = (el.get_attribute("title") or "").strip()
        except Exception:
            t = (el.get_attribute("title") or "").strip()
        code = extract_product_code(t)
        out.append({"text": t, "code": code, "el": el, "frame": fr})
    return out
def _find_latest_analysis_table_nearby(page):
    candidates = []
    for fr in page.frames:
        try:
            scope = fr.locator("xpath=//div[contains(@class,'th-clr-cnt-bottom')]")
            scope = scope if scope.count() else fr
            tds = scope.locator("xpath=.//table[.//td[starts-with(@id,'GUIDE-TextInfoTable-')]]")
            n = tds.count()
            for i in range(n):
                candidates.append((fr, tds.nth(i)))
        except Exception:
            pass
    return candidates[-1] if candidates else (None, None)
def read_text_by_labels(page, wanted_labels):
    fr, tbl = _find_latest_analysis_table_nearby(page)
    if not (fr and tbl and tbl.count()):
        return ""
    row = None
    for t in wanted_labels:
        cand = tbl.locator(
            "xpath=.//tr[td[starts-with(@id,'GUIDE-TextInfoTable-') "
            "and contains(@id,'-TextType') and normalize-space(.)=$t]]"
        ).filter(has_text=t).first
        if cand.count():
            row = cand; break
    if not row:
        for t in wanted_labels:
            low = t.lower()
            cand = tbl.locator(
                "xpath=.//tr[td[starts-with(@id,'GUIDE-TextInfoTable-') "
                "and contains(@id,'-TextType') and "
                f"contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{low}')]]"
            ).first
            if cand.count():
                row = cand; break
    if not row:
        row = tbl.locator("xpath=.//tr[td[starts-with(@id,'GUIDE-TextInfoTable-') and contains(@id,'-Text')]]").first
        if not row.count():
            return ""
    td = row.locator(
        "xpath=.//td[starts-with(@id,'GUIDE-TextInfoTable-') "
        "and contains(@id,'-Text') and not(contains(@id,'-TextType'))]"
    ).first
    if td.count():
        a = td.locator("xpath=.//a[contains(@id,'text_table') and contains(@id,'lines')]").first
        if a.count():
            full = (a.get_attribute('title') or a.get_attribute('aria-label') or '').strip()
            if full:
                return _normalize_text(full)
        txt = _safe_td_text(td)
        if txt:
            return txt
    detail_candidates = fr.locator(
        "xpath=("
        "//textarea[contains(@id,'-Text') and (@readonly or @disabled)] | "
        "//*[@role='textbox' and (not(@contenteditable) or @contenteditable='false')] | "
        "//div[contains(@class,'th-wysi') or contains(@class,'th-txt')][not(@contenteditable) or @contenteditable='false'] | "
        "//div[contains(@class,'text-value') or contains(@class,'TextValue')]"
        ")"
    )
    if detail_candidates.count():
        try:
            return _normalize_text(detail_candidates.first.inner_text())
        except Exception:
            try:
                return _normalize_text(detail_candidates.first.evaluate("n => n.textContent || ''"))
            except Exception:
                pass
    lab = fr.locator("xpath=//*[normalize-space(.)='Text' or contains(normalize-space(.),'Text')]/following::*[1]").first
    if lab.count():
        try:
            return _normalize_text(lab.inner_text())
        except Exception:
            try:
                return _normalize_text(lab.evaluate("n => n.textContent || ''"))
            except Exception:
                pass
    return ""
def read_analysis_summary_for_current_pli(page):
    labels = [
        "Analysis Summary",
        "Analysis/Investigation Summary",
        "Summary of Investigations",
        "Investigation Summary",
        "Analysis/Investigation conclusion",
        "Analysis/Investigation",
    ]
    return (read_text_by_labels(page, labels) or "").strip()
def wait_for_search_with_retries(page, s, *, max_attempts=8, probe_period_ms=2000,
                                 reload_between_attempts=True, total_timeout_ms=240000):
    start = time.time()
    probe_selectors = [s.get('selector')] + (s.get('fallback_selectors', []) or [])
    probe_selectors = [sel for sel in probe_selectors if sel] or ["xpath=//input[contains(@id,'SearchValue')]"]
    def _try_once():
        try:
            loc, ctx, used = wait_find_in_any_frame(page, probe_selectors, timeout_ms=2500, poll_ms=150)
            return (loc, ctx, used)
        except Exception:
            return (None, None, None)
    attempt = 0
    while attempt < max_attempts and (time.time() - start) * 1000 < total_timeout_ms:
        attempt += 1
        log(f"[SSO] probe attempt {attempt}/{max_attempts}")
        loc, ctx, used = _try_once()
        if loc:
            log(f"[SSO] search ready via {used}")
            return (loc, ctx, used)
        page.wait_for_timeout(probe_period_ms)
        if reload_between_attempts:
            try:
                log("[SSO] reloading page to advance SSO…")
                page.reload(wait_until="load")
            except Exception:
                pass
    raise PWTimeout("[SSO] Search not available after retries")
def get_pa_code_to_id(page):
    nav_fr = _find_leftnav_frame(page)
    if not nav_fr:
        any_found = _scan_pa_anywhere(page, "Product Analysis")
        mapping = {}
        for fr, a, txt, did in any_found:
            code = extract_product_code(txt).upper()
            if code and did:
                mapping[code] = did
        log(f"[PA](fallback-section) code→id mapped: {len(mapping)}")
        return mapping
    _ensure_section_expanded(page, "Product Analysis")
    items = _enumerate_pa_items(nav_fr)
    if not items:
        log("[PA] no anchors inside Product Analysis section; scanning section across frames…")
        any_found = _scan_pa_anywhere(page, "Product Analysis")
        mapping = {}
        for fr, a, txt, did in any_found:
            code = extract_product_code(txt).upper()
            if code and did:
                mapping[code] = did
        log(f"[PA](fallback-section) code→id mapped: {len(mapping)}")
        return mapping
    log(f"[PA] detected {len(items)} anchors in Product Analysis section")
    return {(it["code"] or "").upper(): (it["data_id"] or "") for it in items if it["code"] and it["data_id"]}
def click_associated_transactions_tab(page, root_frame):
    return (
        click_tab_by_text(page, root_frame, "Associated Transactions")
        or click_tab_by_text(page, root_frame, "_ovviewset.do_0012")
    )
def _find_assoc_tx_frame(page):
    for fr in page.frames:
        try:
            if fr.locator("xpath=//div[contains(@id,'_Table_bottom') or contains(@id,'_table_bottom')]").first.count():
                return fr
        except Exception:
            pass
    for fr in page.frames:
        try:
            if fr.get_by_role("button", name=re.compile(r"^\s*(Analysis|Investigation)\s*$", re.I)).first.count():
                return fr
            if fr.locator("xpath=//span[contains(@class,'th-bt-span')][.//b[normalize-space(.)='Analysis' or normalize-space(.)='Investigation']]").first.count():
                return fr
        except Exception:
            pass
    return None
def _row_guess_id_type_status(row):
    txid = _row_txid(row)
    status = _row_status_text(row)
    return txid, status
def _collect_unfiltered(fr, header_tbl, body_tbl, scroll_div):
    if not body_tbl or not body_tbl.count():
        return {"product_analysis": [], "investigation": []}
    _scroll_to_load_all_in_div(fr, body_tbl, scroll_div)
    rows = body_tbl.locator("xpath=.//tr[td]")
    pa, inv = [], []
    for i in range(rows.count()):
        row = rows.nth(i)
        txid, gtype, status = _row_guess_id_type_status(row)
        if not txid or not gtype:
            continue
        is_complete = False
        if status:
            s = status.lower()
            is_complete = ("complete" in s) or (s.strip() == "completed") or ("closed - complete" in s)
        else:
            is_complete = False
        if not is_complete:
            continue
        if gtype == "analysis":
            pa.append(txid)
        elif gtype == "investigation":
            inv.append(txid)
    pa = list(dict.fromkeys(pa))
    inv = list(dict.fromkeys(inv))
    log(f"[AssocTx] (unfiltered) PA complete={len(pa)} INV complete={len(inv)}")
    return {"product_analysis": pa, "investigation": inv}
def _assoc_click_filter(fr, label: str) -> bool:
    try:
        btn = fr.get_by_role("button", name=re.compile(rf"^\s*{re.escape(label)}\s*$", re.I)).first
        if btn.count():
            log(f"[AssocTx] clicking filter via role: {label}")
            robust_click(btn, fr)
            try: fr.wait_for_timeout(300)
            except Exception: pass
            return True
    except Exception:
        pass
    try:
        btn = fr.locator(
            "xpath=//span[contains(@class,'th-bt-span')][.//b[normalize-space(.)="
            f"'{label}']]"
        ).first
        if btn.count():
            log(f"[AssocTx] clicking filter via span/b: {label}")
            robust_click(btn, fr)
            try: fr.wait_for_timeout(300)
            except Exception: pass
            return True
    except Exception:
        pass
    try:
        b = fr.locator(f"xpath=//*[normalize-space(.)='{label}']").first
        if b.count():
            cand = b.locator("xpath=ancestor-or-self::*[self::button or self::span or self::a][1]").first
            if cand.count():
                log(f"[AssocTx] clicking filter via generic text: {label}")
                robust_click(cand, fr)
                try: fr.wait_for_timeout(300)
                except Exception: pass
                return True
    except Exception:
        pass
    log(f"[AssocTx] filter button NOT found: {label}")
    return False
def _hdr_indices_from_any(header_tbl, body_tbl):
    labels = []  # list of (index, label_text)
    if header_tbl and header_tbl.count():
        cells = header_tbl.locator("xpath=.//thead//th|.//thead//td|.//tr[1]/*")
        for i in range(cells.count()):
            c = cells.nth(i)
            t = clean(c.inner_text())
            if not t:
                t = (c.get_attribute("aria-label") or c.get_attribute("id") or "").strip()
            labels.append((i, t or ""))
    if not labels:
        first = body_tbl.locator("xpath=.//tr[td]").first
        cells = first.locator("xpath=.//th|.//td")
        for i in range(cells.count()):
            c = cells.nth(i)
            t = (
                c.get_attribute("aria-label")
                or c.get_attribute("headers")
                or c.get_attribute("id")
                or clean(c.inner_text())
                or ""
            ).strip()
            labels.append((i, t))
    def find_idx(patterns):
        for i, t in labels:
            if any(re.search(p, t or "", re.I) for p in patterns):
                return i
        return None
    idx_id = find_idx([r'(?:trans|txn|transaction|work\s*item).*?(?:id|no|number)'])
    idx_type = find_idx([r'(?:trans|txn|transaction|work\s*item|related).*?(?:type|category)'])
    idx_status = find_idx([r'(?:status|state)\b'])
    return idx_id, idx_type, idx_status
def _scroll_to_load_all_in_div(fr, body_tbl, scroll_div):
    if not body_tbl or not body_tbl.count():
        return
    rows = body_tbl.locator("xpath=.//tr[td]")
    target = scroll_div if (scroll_div and scroll_div.count()) else body_tbl
    try:
        box = body_tbl.bounding_box()
        if box:
            fr.mouse.move(box["x"] + box["width"]/2, box["y"] + min(24, box["height"] - 6))
    except Exception:
        pass
    try:
        target.evaluate("n => { n.tabIndex = 0; n.focus(); }")
    except Exception:
        pass
    prev = -1
    stagnant = 0
    for _ in range(240):  # ~30s worst-case
        n = rows.count()
        log(f"[AssocTx] rows visible: {n}")
        if n == prev:
            stagnant += 1
            if stagnant >= 6:
                break
        else:
            stagnant = 0
        prev = n
        try:
            if n > 0:
                rows.nth(n - 1).scroll_into_view_if_needed(timeout=500)
        except Exception:
            pass
        try:
            target.press("PageDown")
        except Exception:
            try: fr.keyboard.press("PageDown")
            except Exception: pass
        try:
            target.press("End")
        except Exception:
            try: fr.keyboard.press("End")
            except Exception: pass
        try:
            fr.mouse.wheel(0, 1800)
        except Exception:
            pass
        fr.wait_for_timeout(140)
_TRANS_HEADER_RX = re.compile(
    r"(transaction|work\s*item|related|type|category|status|state|number|id|no\b|ref|reference)",
    re.I,
)
_PLI_HEADER_RX = re.compile(
    r"(product|description|lot|sn|serial|qty|quantity|uom|unit)", re.I
)
def _table_header_labels(tbl):
    labels = []
    cells = tbl.locator("xpath=.//thead//th|.//thead//td|.//tr[1]/*")
    for j in range(cells.count()):
        c = cells.nth(j)
        t = clean(c.inner_text()) or (c.get_attribute("aria-label") or c.get_attribute("id") or "")
        t = (t or "").strip()
        if t:
            labels.append(t)
    return labels

def _score_header_labels(labels):
    if not labels:
        return -999
    trans_hits = sum(1 for t in labels if _TRANS_HEADER_RX.search(t))
    pli_hits   = sum(1 for t in labels if _PLI_HEADER_RX.search(t))
    return (trans_hits * 3) - (pli_hits * 4)
def _pick_assoc_grid_table(fr):
    best = (-9999, None, None, None, [])
    bodies = fr.locator(
        "xpath=//div[contains(@id,'_Table_bottom') or contains(@id,'_table_bottom')]"
        "//table[contains(@class,'th-clr-table')]"
    )
    for i in range(bodies.count()):
        body = bodies.nth(i)
        header = body.locator(
            "xpath=ancestor::div[1]/preceding-sibling::div[1]//table[contains(@class,'th-clr-table')]"
        ).first
        labels = _table_header_labels(header if header.count() else body)
        score = _score_header_labels(labels)
        if _is_bad_table(body, labels):
            score -= 5000
        sigs = _table_id_signatures(body)
        if any(_ASSOC_CELL_ID_RX.search(s) for s in sigs):
            score += 8000
        labset = set(labels)
        if "Transaction ID" in labset:
            score += 500
        if "Status" in labset:
            score += 300
        sc = body.locator("xpath=ancestor::div[contains(@style,'overflow')][1]").first
        if sc.count():
            score += 50
        if score > best[0]:
            best = (score, header if header.count() else None, body, sc if sc.count() else None, labels)
    if best[1] is None and best[2] is None:
        any_tbls = fr.locator("xpath=//table[contains(@class,'th-clr-table')]")
        for i in range(any_tbls.count()):
            tb = any_tbls.nth(i)
            labels = _table_header_labels(tb)
            score = _score_header_labels(labels)
            if _is_bad_table(tb, labels): score -= 5000
            sigs = _table_id_signatures(tb)
            if any(_ASSOC_CELL_ID_RX.search(s) for s in sigs): score += 8000
            if score > best[0]:
                best = (score, None, tb, None, labels)
    _, header_tbl, body_tbl, scroll_div, headers = best
    if body_tbl:
        log(f"[AssocTx] chose grid with headers: {headers}")
        if _is_bad_table(body_tbl, headers) and not _is_assoc_tx_table(headers, body_tbl):
            log("[AssocTx] rejecting non-transaction grid (attachments/text/partners)")
            return None, None, None, []
    else:
        log("[AssocTx] no suitable transaction grid found in this frame")
    return header_tbl, body_tbl, scroll_div, headers
def _node_signature(loc):
    try:
        return loc.evaluate("n => (n.innerText || '').slice(0, 1200)") or ""
    except Exception:
        try:
            return clean(loc.inner_text())[:1200]
        except Exception:
            return ""
_ASSOC_CELL_ID_RX = re.compile(r'\bGUIDE-AssociatedTransactionsTable-', re.I)
_BAD_TABLE_ID_RXS = [
    re.compile(r'\bGUIDE-AttachmentsTable-', re.I),
    re.compile(r'\bGUIDE-TextInfoTable-', re.I),
    re.compile(r'\bGUIDE-PartnersTable-', re.I),
    re.compile(r'\bGUIDE-ProductLineItemsTable-', re.I),
]
def _table_id_signatures(tbl):
    sigs = []
    try:
        cells = tbl.locator("xpath=.//th|.//td")
        for j in range(min(cells.count(), 200)):
            sid = (cells.nth(j).get_attribute("id") or "")
            if sid:
                sigs.append(sid)
    except Exception:
        pass
    return sigs
def _is_assoc_tx_table(headers, body_tbl):
    labels = [h.lower() for h in (headers or [])]
    if ("transaction id" in labels and "status" in labels):
        return True
    sigs = _table_id_signatures(body_tbl)
    if any(_ASSOC_CELL_ID_RX.search(s) for s in sigs):
        return True
    return False
def _is_bad_table(body_tbl, headers):
    sigs = _table_id_signatures(body_tbl)
    if any(rx.search(s) for rx in _BAD_TABLE_ID_RXS for s in sigs):
        return True
    h = set((headers or []))
    if {"Name","Document Type","Folder Path"}.issubset(h):
        return True  # attachments
    return False
def read_associated_transactions_complete(page, root_frame):
    click_associated_transactions_tab(page, root_frame)
    fr = None
    for _ in range(20):
        fr = _find_assoc_tx_frame(page)
        if fr:
            break
        page.wait_for_timeout(250)
    if not fr:
        log("[AssocTx] grid frame not found")
        return {"product_analysis": [], "investigation": []}
    def _collect_for(label):
        h0, b0, sc0, hdr0 = _pick_assoc_grid_table(fr)
        sig_before = _node_signature(b0) if b0 else ""
        if not _assoc_click_filter(fr, label):
            return []
        header_tbl = body_tbl = scroll_div = None
        headers = []
        changed = False
        for _ in range(40):  # ~8s
            page.wait_for_timeout(200)
            h1, b1, sc1, hdr1 = _pick_assoc_grid_table(fr)
            if b1 and _is_assoc_tx_table(hdr1, b1):
                sig_after = _node_signature(b1)
                header_tbl, body_tbl, scroll_div, headers = h1, b1, sc1, hdr1
                if (sig_after and sig_after != sig_before) or not sig_before:
                    changed = True
                    break
        if not body_tbl:
            log(f"[AssocTx] no AssociatedTransactions grid after clicking {label} — skipping")
            return []
        _scroll_to_load_all_in_div(fr, body_tbl, scroll_div)
        rows = body_tbl.locator("xpath=.//tr[td]")
        idx_id, idx_type, idx_status = _hdr_indices_from_any(header_tbl, body_tbl)
        ids_complete, ids_any = [], []
        n = rows.count()
        for i in range(n):
            row = rows.nth(i)
            cells = row.locator("xpath=.//th|.//td")
            txid = None
            status = None
            if idx_id is not None and cells.count() > idx_id:
                txid_txt = clean(cells.nth(idx_id).inner_text())
                txid = _find_first_match([r"\b\d{5,}\b", r"\b[A-Z]{2,5}[-_ ]?\d{4,}\b"], txid_txt) or txid_txt
            else:
                txid = _row_txid(row)

            if idx_status is not None and cells.count() > idx_status:
                status = clean(cells.nth(idx_status).inner_text())
            else:
                status = _row_status_text(row)

            dbg = f"id={txid or '-'} status={(status or '').strip()!r}"
            log(f"[AssocTx:{label}] row {i+1}/{n}: {dbg}")

            if not txid:
                continue
            ids_any.append(txid)
            if _row_is_complete(row) or re.search(r'\bcomplete(d)?\b', (status or ''), re.I):
                ids_complete.append(txid)
        ids_complete = list(dict.fromkeys(ids_complete))
        ids_any = list(dict.fromkeys(ids_any))
        if ids_complete:
            return ids_complete
        if ids_any:
            log(f"[AssocTx:{label}] no explicit 'Complete' statuses found; returning all ({len(ids_any)})")
            return ids_any
        return []
    has_btns = (
        fr.get_by_role("button", name=re.compile(r"^\s*(Analysis|Investigation)\s*$", re.I)).first.count() or
        fr.locator("xpath=//span[contains(@class,'th-bt-span')][.//b[normalize-space(.)='Analysis' or normalize-space(.)='Investigation']]").first.count()
    )
    if not has_btns:
        log("[AssocTx] filter buttons not present in detected frame; falling back to unfiltered parse")
        header_tbl, body_tbl, scroll_div, _ = _pick_assoc_grid_table(fr)
        if not body_tbl:
            return {"product_analysis": [], "investigation": []}
        return _collect_unfiltered(fr, header_tbl, body_tbl, scroll_div)
    pa = _collect_for("Analysis")
    inv = _collect_for("Investigation")
    log(f"[AssocTx] Product Analysis (Complete or fallback): {pa}")
    log(f"[AssocTx] Investigations (Complete or fallback): {inv}")
    return {"product_analysis": pa, "investigation": inv}
def summary_has_product_id(text: str, product_id: str) -> bool:
    if not text or not product_id:
        return False
    pid = str(product_id).strip()
    return re.search(rf"\b{re.escape(pid)}\b", text, re.IGNORECASE) is not None
def collect_product_analysis(page, root_frame, known_products):
    id_by_code = {}
    wanted_codes = set()
    for p in known_products:
        code = (p.get("code") or extract_product_code(p.get("desc",""))).upper()
        if code:
            wanted_codes.add(code)
            pid = (p.get("id") or "").strip()
            if pid:
                id_by_code[code] = pid
    nav_clicked = click_left_nav_product_analysis(page)
    if not nav_clicked:
        default_msg = "Information provided to Medtronic indicated that the complaint device was not available for evaluation."
        return {(p.get("id") or p.get("code") or "").strip(): default_msg
                for p in known_products if (p.get("id") or p.get("code"))}
    nav_fr = _find_leftnav_frame(page)
    if not nav_fr:
        nav_items = _scan_pa_anywhere(page, "Product Analysis")
        items = [{"i": i, "text": txt, "code": extract_product_code(txt), "data_id": did, "el": a, "frame": fr}
                 for i, (fr, a, txt, did) in enumerate(nav_items)]
    else:
        _ensure_section_expanded(page, "Product Analysis")
        items = _enumerate_pa_items(nav_fr)
        if not items:
            log("[PA] no anchors in-section; scanning section across frames…")
            nav_items = _scan_pa_anywhere(page, "Product Analysis")
            items = [{"i": i, "text": txt, "code": extract_product_code(txt), "data_id": did, "el": a, "frame": fr}
                     for i, (fr, a, txt, did) in enumerate(nav_items)]
    items = [it for it in items if (it.get("code","").upper() in wanted_codes)]
    log("[PA] anchors (Product Analysis only): " + "; ".join([repr(it["text"]) for it in items[:10]]))
    summaries_by_text = {}
    for it in items:
        anchors_now = _product_analysis_anchor_locator(nav_fr) if nav_fr else None
        target = anchors_now.nth(it["i"]) if (anchors_now and anchors_now.count() > it["i"]) else None
        if (not target or not target.count()) and it.get("data_id") and nav_fr:
            target = _pa_anchor_by_data_id(nav_fr, it["data_id"])
        if not target or not target.count():
            target = it.get("el")
        if not target or not target.count():
            continue
        content_fr = _content_frame(page)
        prev_sig = _textinfo_signature(page)
        if not robust_click_plus(target, it.get("frame") or nav_fr or page.main_frame):
            log(f"[PA] click failed for: {it['text']!r} data_id={it.get('data_id','')}")
            continue
        changed = wait_for_textinfo_change(page, prev_sig, timeout=10000)
        if not changed:
            try:
                content_fr.wait_for_selector(
                    "xpath=//div[contains(@class,'th-clr-cnt-bottom')]"
                    "//td[starts-with(@id,'GUIDE-TextInfoTable-') and contains(@id,'-TextType')]",
                    timeout=3000
                )
                changed = True
            except Exception:
                pass
        if not changed:
            content_fr.wait_for_timeout(350)
        txt = read_analysis_summary_for_current_pli(page).strip()
        expected_pid = id_by_code.get((it.get("code") or "").upper(), "")
        if txt and expected_pid and not summary_has_product_id(txt, expected_pid):
            log(f"[PA] discard summary for {it['text']!r}: missing product id {expected_pid}")
            txt = ""
        log(f"[PA] read summary for {it['text']!r}: {('len='+str(len(txt)) if txt else 'EMPTY')}")
        if txt:
            summaries_by_text[it["text"]] = txt
    default_msg = "Information provided to Medtronic indicated that the complaint device was not available for evaluation."
    results = {}
    for p in known_products:
        pid  = (p.get("id") or "").strip()
        code = (p.get("code") or extract_product_code(p.get("desc",""))).strip().upper()
        best = ""
        if code:
            for t, s in summaries_by_text.items():
                if code == (extract_product_code(t) or "").upper():
                    best = s
                    break
        key = code or pid or p.get("desc","")
        results[key] = best or default_msg
    return results
def robust_click(el, frame, timeout_ms=8000):
    try:
        el.scroll_into_view_if_needed(timeout=2000)
    except Exception:
        pass
    try:
        el.click()
        return True
    except Exception:
        pass
    try:
        el.evaluate("e => { e.scrollIntoView({block:'center'}); e.click(); }")
        return True
    except Exception:
        pass
    try:
        box = el.bounding_box()
        if box:
            frame.mouse.click(box["x"] + min(5, box["width"]/2), box["y"] + min(5, box["height"]/2))
            return True
    except Exception:
        pass
    return False
def expand_full_text_if_collapsed(frame):
    try:
        link = frame.locator("xpath=//a[contains(@id,'text_table') and contains(@id,'lines')]").first
        if link.count():
            robust_click(link, frame)
            frame.wait_for_timeout(200)
    except Exception:
        pass
def extract_product_code(desc: str) -> str:
    m = re.search(r'[A-Z0-9_]+', (desc or '').upper())
    return m.group(0) if m else ''
def active_content_frames(page, cfg):
    name_rx = cfg.get('content_frame_name_regex') or cfg.get('frame_name_regex')
    url_rx  = cfg.get('frame_url_regex')
    frames = []
    for fr in page.frames:
        try:
            nm, url = (fr.name or ""), (fr.url or "")
        except Exception:
            nm, url = "", ""
        if (name_rx and re.search(name_rx, nm, re.I)) or (url_rx and re.search(url_rx, url, re.I)):
            frames.append(fr)
    if frames:
        return frames
    for fr in page.frames:
        if (fr.name or "") == "WorkAreaFrame1":
            return [fr]
    return list(page.frames)
def _find_latest_text_table_in(fr):
    scope = fr.locator("xpath=//div[contains(@class,'th-clr-cnt-bottom')]")
    scope = scope if scope.count() else fr
    tds = scope.locator("xpath=.//table[.//td[starts-with(@id,'GUIDE-TextInfoTable-')]]")
    return tds.nth(tds.count()-1) if tds.count() else None
def _suppress_clicks_enable(frame):
    try:
        frame.evaluate("""
            window.__mdt_suppress = e => { e.stopPropagation(); e.preventDefault(); };
            document.addEventListener('mousedown', window.__mdt_suppress, true);
            document.addEventListener('click', window.__mdt_suppress, true);
            document.addEventListener('mouseup', window.__mdt_suppress, true);
        """)
    except Exception:
        pass
def _suppress_clicks_disable(frame):
    try:
        frame.evaluate("""
            if (window.__mdt_suppress) {
                document.removeEventListener('mousedown', window.__mdt_suppress, true);
                document.removeEventListener('click', window.__mdt_suppress, true);
                document.removeEventListener('mouseup', window.__mdt_suppress, true);
                window.__mdt_suppress = null;
            }
        """)
    except Exception:
        pass
def _normalize_text(s: str) -> str:
    s = (s or "").replace("\xa0", " ")
    s = re.sub(r'\r?\n\s*\r?\n+', '\n\n', s)
    s = re.sub(r'[ \t]+', ' ', s)
    s = re.sub(r'\s*\n\s*', '\n', s).strip()
    return s
def _safe_td_text(td):
    if not td or not td.count():
        return ""
    try:
        raw = td.evaluate("""
            (node) => {
                // prefer the expander link’s attributes (full content)
                const a = node.querySelector("a[id*='text_table'][id*='lines']");
                if (a) {
                    const t = a.getAttribute('title') || a.getAttribute('aria-label');
                    if (t && t.trim()) return t;
                }
                return node.textContent || '';
            }
        """) or ""
    except Exception:
        try:
            raw = td.evaluate("n => n.textContent || ''") or ""
        except Exception:
            raw = ""
    return _normalize_text(raw)
def _pa_try_expand(fr):
    header = fr.locator("xpath=//*[contains(@class,'left-nav')]//*[contains(@class,'ProductAnalysis')]").first
    if not header.count(): 
        return False
    clicker = header.locator("xpath=following-sibling::*[contains(@class,'clicker')][1]").first
    container = header.locator("xpath=following-sibling::*[contains(@class,'data-wrapper')][1]").first
    if container.count():
        try:
            is_hidden = container.evaluate("n => getComputedStyle(n).display === 'none'")
        except Exception:
            is_hidden = False
        if not is_hidden:
            return True
    if clicker.count():
        try:
            robust_click(clicker, fr)
            fr.wait_for_timeout(250)
        except Exception:
            pass
    if container.count():
        try:
            return not container.evaluate("n => getComputedStyle(n).display === 'none'")
        except Exception:
            return True
    return False
def _pa_anchor_by_data_id(fr, data_id: str):
    if not data_id:
        return None
    return fr.locator(
        "xpath=//*[contains(@class,'left-nav')]"
        "//div[contains(@class,'ProductAnalysis')]"
        "/following-sibling::div[contains(@class,'clicker')][1]"
        "/following-sibling::div[contains(@class,'data-wrapper')][1]"
        f"//a[contains(@class,'GUIDE-sideNav') and @data-trans-id='{data_id}']"
    ).first
def robust_click_plus(el, frame):
    try:
        el.scroll_into_view_if_needed(timeout=1500)
    except Exception:
        pass
    try:
        el.click(timeout=3000, force=True)
        return True
    except Exception:
        pass
    try:
        el.evaluate("""
            e => {
                const ev1 = new MouseEvent('mousedown', {bubbles:true, cancelable:true});
                const ev2 = new MouseEvent('mouseup',   {bubbles:true, cancelable:true});
                const ev3 = new MouseEvent('click',     {bubbles:true, cancelable:true});
                e.dispatchEvent(ev1); e.dispatchEvent(ev2); e.dispatchEvent(ev3);
            }
        """)
        return True
    except Exception:
        pass
    try:
        el.evaluate("e => { e.scrollIntoView({block:'center'}); e.click(); }")
        return True
    except Exception:
        pass
    try:
        box = el.bounding_box()
        if box:
            frame.mouse.click(box["x"] + min(5, box["width"]/2), box["y"] + min(5, box["height"]/2))
            return True
    except Exception:
        pass
    return False
def main():
    if len(sys.argv) < 3:
        print("Usage: python scrape_and_generate.py <complaint_id> <config.yaml>")
        sys.exit(2)
    complaint_id = sys.argv[1]
    cfg_path = Path(sys.argv[2])
    cfg = yaml.safe_load(cfg_path.read_text())
    template_path = Path(cfg['template_path']).expanduser()
    out_dir = Path(cfg.get('output_dir', '.')).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=cfg.get('headless', False))
        context = browser.new_context()
        page = context.new_page()
        log(f"Navigating to CRM: {cfg['crm_url']}")
        page.goto(cfg['crm_url'], wait_until="load")
        sso_wait = cfg.get('sso_pause_seconds', 0)
        if sso_wait > 0:
            s = cfg.get('search', {})
            try:
                wait_find_in_any_frame(page, [s.get('selector')] + (s.get('fallback_selectors', []) or []),
                                    timeout_ms=2500, poll_ms=150)
                print("[SSO] Search is already available; skipping SSO wait.")
            except Exception:
                print(f"[SSO] Search not ready; running extended SSO retries (up to ~{cfg.get('sso_total_timeout_ms', 240000)//1000}s)…")
                try:
                    wait_for_search_with_retries(
                        page, s,
                        max_attempts=cfg.get('sso_max_attempts', 8),
                        probe_period_ms=cfg.get('sso_probe_period_ms', 2000),
                        reload_between_attempts=cfg.get('sso_reload_between_attempts', True),
                        total_timeout_ms=cfg.get('sso_total_timeout_ms', 240000),
                    )
                except Exception:
                    page.wait_for_timeout(sso_wait * 1000)
        frame = find_app_frame(
            page,
            frame_name_regex=cfg.get('frame_name_regex'),
            url_regex=cfg.get('frame_url_regex')
        )
        if 'search' in cfg:
            s = cfg['search']
            if s.get('selector'):
                try:
                    fallbacks = s.get('fallback_selectors', [])
                    extra_defaults = [
                        "xpath=//input[@id='C12_W37_V38_SearchValue']",
                        "css=#C1_W1_V2_C9_W28_V29_C12_W37_V38_launcher\\.do > span > table > tbody > tr > td > span.th-if-wrapper > input",
                        "xpath=/html/body/form/div[5]/div/table/tbody/tr[1]/td/div/div/div/div/table/tbody/tr/td[1]/div/div/span/table/tbody/tr/td/span[3]/input",
                    ]
                    all_selectors = [s['selector']] + [sel for sel in fallbacks + extra_defaults if sel not in fallbacks]
                    log("[search] polling every 2s for search input…")
                    poll_deadline = time.time() + s.get('pre_wait_timeout', 30000)/1000.0
                    target = target_ctx = used_sel = None
                    while time.time() < poll_deadline:
                        try:
                            target, target_ctx, used_sel = wait_find_in_any_frame(page, all_selectors, timeout_ms=1, poll_ms=1)
                            break
                        except Exception:
                            log("[search] not visible yet; sleeping 2s")
                            time.sleep(2)
                    if not target:
                        target, target_ctx, used_sel = wait_find_in_any_frame(page, all_selectors, timeout_ms=5000)
                    log(f"[search] Found input via selector: {used_sel} in frame url={getattr(target_ctx, 'url', '')} name={getattr(target_ctx, 'name', '')}")
                    if s.get('clear', True):
                        try:
                            target.fill("")
                        except Exception:
                            pass
                    target.click()
                    frame.wait_for_timeout(100)
                    try:
                        log(f"[search] filling complaint id: {complaint_id}")
                        target.fill(complaint_id)
                    except:
                        target.type(complaint_id, delay=30)
                    try:
                        target.evaluate("el => { el.dispatchEvent(new Event('input',{bubbles:true})); el.blur(); el.focus(); }")
                    except: pass
                    submitted = False
                    if s.get('submit_selector'):
                        try:
                            btn = target_ctx.locator(s['submit_selector']).first
                            if btn.count():
                                log("[search] clicking submit")
                                btn.click()
                                submitted = True
                        except Exception:
                            pass
                    if not submitted:
                        key = s.get('press_key', 'Enter')
                        try:
                            log(f"[search] pressing key: {key}")
                            target.press(key)
                        except Exception: pass
                    if s.get('wait_for'):
                        log(f"[wait] waiting for results: {s['wait_for']}")
                        target_ctx.wait_for_selector(s['wait_for'], timeout=s.get('wait_timeout', 60000))
                    else:
                        target_ctx.wait_for_load_state("networkidle")
                        target_ctx.wait_for_timeout(s.get('post_wait_ms', 3000))
                except Exception as e:
                    log(f"[ERROR] Search failed: {e}")
                    print(f"[Search] Failed to drive search: {e}")
                    dump_frames_debug(page, basename='debug')
                    try:
                        page.screenshot(path="debug_search_failure.png", full_page=True)
                        print("Saved debug_search_failure.png")
                        log("Saved debug_search_failure.png")
                    except Exception:
                        pass
        values = {}
        if not values.get("todays_date"):
            values["todays_date"] = date.today().isoformat()
        values['complaint_id'] = complaint_id
        for key, conf in cfg.get('field_map', {}).items():
            if isinstance(conf, str):
                values[key] = get_by_label(frame, conf)
            elif isinstance(conf, list):
                values[key] = get_by_label(frame, conf)
            elif isinstance(conf, dict):
                if conf.get('type') == 'label':
                    values[key] = get_by_label(frame, conf.get('labels', []))
                elif conf.get('type') == 'grid':
                    grid_xpath = conf['grid_xpath']
                    row = conf['row']
                    col = conf['col']
                    values[key] = get_grid_row_col(frame, grid_xpath, row, col)
                elif conf.get('type') == 'literal':
                    values[key] = conf.get('value', '')
            else:
                values[key] = ""
        for k, v in cfg.get('defaults', {}).items():
            values.setdefault(k, v)
        try:
            log("[step 1] Partners tab → IR name & facility")
            if click_partners_tab(page, frame):
                pframe = find_partners_frame(page)
                if not pframe:
                    debug_frames_for_partners(page)
                    print("[Partners] Could not locate the partners frame.")
                    log("[Partners] Could not locate the partners frame.")
                else:
                    _debug_list_pf_from_correct_table(pframe)
                    irname = get_initial_reporter_name(pframe)
                    if irname:
                        values['ir_name'] = irname
                    facility_block = get_facility_name_and_address(pframe)
                    if facility_block:
                        values['ir_with_address'] = facility_block
                    print("[Partners] ir_name =", values.get('ir_name', ''))
                    print("[Partners] ir_with_address =", values.get('ir_with_address', ''))
                    log(f"[Partners] ir_name = {values.get('ir_name','')}")
                    log(f"[Partners] ir_with_address = {values.get('ir_with_address','')}")
            else:
                print("[Partners] Could not open Partners tab; leaving ir_* fields from label map/fallbacks.")
                log("[Partners] Could not open Partners tab; leaving ir_* fields from label map/fallbacks.")
        except Exception as e:
            print(f"[Partners] Error scraping Partners tab: {e}")
            log(f"[ERROR] Partners scrape: {e}")
        log("[step 2] Additional External References → rb_reference & report_number")
        ext = read_external_refs(page, frame)
        if ext.get("rb_reference"):
            values["rb_reference"] = ext["rb_reference"]
        if ext.get("report_number"):
            values["report_number"] = ext["report_number"]
        log(f"[AER] rb_reference={values.get('rb_reference','')}, report_number={values.get('report_number','')}")
        log("[step 3] Dates tab → event_date")
        event_date_text = get_event_date(page)
        if event_date_text:
            values['event_date'] = event_date_text
        log(f"[Dates] event_date={values.get('event_date','')}")
        log("[step 4] Product Line Items → all rows")
        products = read_all_products(page, frame)
        log(f"[PLI] rows detected: {len(products)}")
        log("[step 5] Text Info → event_description")
        prev_sig = _textinfo_signature(page)
        desc = read_event_description(page, frame)
        if not desc and wait_for_textinfo_change(page, prev_sig, timeout=8000):
            desc = read_event_description(page, frame)
        if desc:
            values["event_description"] = desc
        log(f"[Text] description length: {len(values.get('event_description',''))}")
        log("[step 6] Associated Transactions → collect Complete Investigation/Product Analysis IDs")
        assoc = read_associated_transactions_complete(page, frame)
        print(f"[AssociatedTx] Product Analysis (Complete): {assoc['product_analysis']}")
        print(f"[AssociatedTx] Investigations (Complete): {assoc['investigation']}")
        for idx, p in enumerate(products[:3], start=1):
            code = (p.get("code") or extract_product_code(p.get("desc",""))).upper()
            values[f"product_id_{idx}"] = (p.get("id") or code)
            values[f"product_desc_{idx}"]  = p.get("desc","")
            values[f"product_sn_{idx}"]    = p.get("sn","")
            values[f"product_lot_{idx}"]   = p.get("lot","")
            values[f"serial_or_lot_{idx}"] = " / ".join([s for s in [p.get('sn',''), p.get('lot','')] if s])
            values["assoc_tx_product_analysis_ids"] = ", ".join(assoc["product_analysis"])
            values["assoc_tx_investigation_ids"]    = ", ".join(assoc["investigation"])
        if len(products) > 3:
            extras = [f"{p['id']} — {p['desc']}" for p in products[3:]]
            values.setdefault("product_extras", "\n".join(extras))
        log("[step 7] Investigation side panel → Summary of Investigations per item")
        if click_left_nav_investigation(page):
            inv_items = list_side_nav_items(page, "Investigations")
            inv_lines = []
            for it in inv_items:
                if not robust_click(it["el"], it["frame"]):
                    continue
                text = read_text_by_labels(page, [
                    "Summary of Investigations",
                    "Investigation Summary",
                    "Analysis/Investigation Summary"
                ])
                if text:
                    inv_lines.append(f"{it['text']} — {text}")
            if inv_lines:
                values.setdefault("investigation_summary", "\n\n".join(inv_lines))
        else:
            log("[Investigation] Left nav not found; skipping.")
        log("Collected fields:")
        log(json.dumps(values, indent=2))    
        out_name = cfg.get('output_name_pattern', 'Customer_Letter_{complaint_id}.docx').format(**values)
        out_path = out_dir / out_name
        fill_docx(str(template_path), str(out_path), values)
        log(f"Generated: {out_path}")
        browser.close()
if __name__ == "__main__":
    main()