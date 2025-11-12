import re, sys, time, json
from pathlib import Path
from datetime import date
import yaml
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from docx import Document
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
def wait_find_in_any_frame(page, selectors, timeout_ms=30000, poll_ms=300):
    import time
    deadline = time.time() + (timeout_ms/1000.0)
    tried = set()
    while time.time() < deadline:
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
                        return loc, fr, sel
                except Exception:
                    continue
        time.sleep(poll_ms/1000.0)
    raise PWTimeout(f"Could not find element in any frame for selectors: {selectors}")
def replace_everywhere(doc, mapping):
    for p in doc.paragraphs:
        text = p.text
        for k, v in mapping.items():
            text = text.replace(f"{{{{{k}}}}}", v or "")
        if text != p.text:
            for r in p.runs:
                r.clear()
            p.text = text
    for t in doc.tables:
        for r in t.rows:
            for c in t.rows[0].cells:  # no-op: just keeps structure similar
                pass
        for r in t.rows:
            for c in r.cells:
                text = c.text
                for k, v in mapping.items():
                    text = text.replace(f"{{{{{k}}}}}", v or "")
                c.text = text
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
def _row_by_partner_function(frame, names):
    if isinstance(names, str):
        names = [names]
    for fn in names:
        tr = frame.locator(
            f"xpath=//td[@aria-label='Partner Function' and normalize-space(.)='{fn}']/ancestor::tr[1]"
        ).first
        if tr.count():
            return tr
    for fn in names:
        tr = frame.locator(
            f"xpath=//tr[td and normalize-space(td[1])='{fn}']"
        ).first
        if tr.count():
            return tr
    for fn in names:
        low = fn.lower()
        tr = frame.locator(
            "xpath=//td[@aria-label='Partner Function' and "
            f"contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{low}')]/ancestor::tr[1]"
        ).first
        if tr.count():
            return tr
    for fn in names:
        low = fn.lower()
        tr = frame.locator(
            "xpath=//tr[td and contains(translate(normalize-space(td[1]),"
            f" 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{low}')]"
        ).first
        if tr.count():
            return tr
    return None
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
def _cell_in_row_by_aria_fuzzy(tr_loc, logical_name):
    candidates = []
    if logical_name.lower() == "name":
        candidates = [
            "xpath=.//td[@aria-label='Name']",
            "xpath=.//td[contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'name')]",
        ]
    elif logical_name.lower() == "address":
        candidates = [
            "xpath=.//td[@aria-label='Address']",
            "xpath=.//td[@aria-label='address_short']",
            "xpath=.//td[contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'address')]",
            "xpath=.//td[contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'addr')]",
        ]
    else:
        candidates = [f"xpath=.//td[@aria-label='{logical_name}']"]
    for sel in candidates:
        td = tr_loc.locator(sel).first
        if td.count():
            return clean(td.inner_text())
    return ""
def _debug_list_partner_functions(frame):
    try:
        pfs = frame.locator("xpath=//td[@aria-label='Partner Function']")
        n = pfs.count()
        if n == 0:
            pfs = frame.locator("xpath=//tr[td]/td[1]")
            n = pfs.count()
        print(f"[Partners] Found {n} partner rows")
        for i in range(min(n, 30)):
            print("  -", clean(pfs.nth(i).inner_text()))
    except Exception as e:
        print("[Partners] PF debug error:", e)
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
                    loc.click()
                    return fr
            except Exception:
                pass
    return None
def _pli_table(frame):
    return frame.locator("xpath=//table[.//td[starts-with(@id,'GUIDE-ProductLineItemsTable-')]]").first
def read_all_products(page, root_frame):
    click_tab_by_text(page, root_frame, "Product Line Items") or \
    click_tab_by_text(page, root_frame, "_ovviewset.do_0002")
    fr = find_frame_with(page, "xpath=//td[starts-with(@id,'GUIDE-ProductLineItemsTable-')]")
    if not fr:
        return []
    tbl = _pli_table(fr)
    if not tbl or not tbl.count():
        return []
    rows = tbl.locator(
        "xpath=.//tr[td[starts-with(@id,'GUIDE-ProductLineItemsTable-') and contains(@id,'-Product')] ]"
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
            "xpath=.//td[starts-with(@id,'GUIDE-ProductLineItemsTable-') and contains(@id,'-Lot')]",     # covers -Lot and -LotVal
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
    return out
def _dates_table(frame):
    return frame.locator("xpath=//table[.//td[starts-with(@id,'GUIDE-DatesTable')]]").first
def get_event_date(page):
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
def _expand_if_truncated(frame):
    try:
        link = frame.locator("xpath=//a[contains(@id,'text_table') and contains(@id,'lines')]").first
        if link.count():
            try:
                link.scroll_into_view_if_needed(timeout=1000)
            except Exception:
                pass
            try:
                link.click()
            except Exception:
                link.evaluate("e => e.click()")
            frame.wait_for_timeout(200)
    except Exception:
        pass
def read_event_description(page, root_frame):
    click_tab_by_text(page, root_frame, "Text Info") or \
    click_tab_by_text(page, root_frame, "_ovviewset.do_0006")
    try:
        root_frame.wait_for_selector("xpath=//td[starts-with(@id,'GUIDE-TextInfoTable-')]", timeout=8000)
    except Exception:
        pass
    fr, tbl = _find_latest_text_table(page)
    if not (fr and tbl and tbl.count()):
        return ""
    _expand_if_truncated(fr)
    want_types = [
        "Incident description", "Incident description / Reason for report",
        "Reason for report", "Description of Event", "Event Description",
        "Narrative", "HCP Narrative"
    ]
    row = None
    for t in want_types:
        cand = tbl.locator(
            "xpath=.//tr[td[starts-with(@id,'GUIDE-TextInfoTable-') and contains(@id,'-TextType') and normalize-space(.)=$t]]"
        ).filter(has_text=t).first
        if cand.count():
            row = cand
            break
    if not row or not row.count():
        for t in want_types:
            low = t.lower()
            cand = tbl.locator(
                "xpath=.//tr[td[starts-with(@id,'GUIDE-TextInfoTable-') and contains(@id,'-TextType') "
                f"and contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{low}')]]"
            ).first
            if cand.count():
                row = cand
                break
    if not row or not row.count():
        row = tbl.locator("xpath=.//tr[td[starts-with(@id,'GUIDE-TextInfoTable-') and contains(@id,'-Text')]]").first
        if not row.count():
            return ""
    text_td = row.locator(
        "xpath=.//td[starts-with(@id,'GUIDE-TextInfoTable-') and contains(@id,'-Text') and not(contains(@id,'-TextType'))]"
    ).first
    if not text_td.count():
        return ""
    raw = text_td.evaluate("el => el.textContent || ''")
    para = re.sub(r'\r?\n\s*\r?\n+', '\n\n', raw)
    para = re.sub(r'[ \t\xa0]+', ' ', para)
    para = re.sub(r'\s*\n\s*', '\n', para).strip()
    return para
def click_left_nav_product_analysis(page):
    candidates = [
        "xpath=//div[contains(@class,'left-nav')]//div[normalize-space(.)='ProductAnalysis']",
        "xpath=//div[contains(@class,'left-nav')]//div[normalize-space(.)='Product Analysis']",
        "xpath=//*[contains(@class,'left-nav')]//*[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'product analysis')]",
    ]
    for sel in candidates:
        for fr in page.frames:
            try:
                loc = fr.locator(sel).first
                if loc.count():
                    loc.click()
                    try:
                        fr.wait_for_selector("xpath=//a[contains(@class,'GUIDE-sideNav')]", timeout=4000)
                    except Exception:
                        pass
                    return True
            except Exception:
                continue
    return False
def _find_frame_with_selector(page, sel):
    for fr in page.frames:
        try:
            if fr.locator(sel).first.count():
                return fr
        except Exception:
            pass
    return None
def _textinfo_signature(page):
    fr, tbl = _find_latest_analysis_table_nearby(page)  # or _find_latest_text_table for Text Info
    if not (fr and tbl and tbl.count()):
        return ""
    try:
        return tbl.evaluate("t => (t.innerText || '').slice(0, 300)")
    except Exception:
        try:
            return tbl.inner_text()[:300]
        except Exception:
            return ""
def wait_for_textinfo_change(page, previous_sig, timeout=10000):
    start = time.time()
    while time.time() - start < (timeout/1000.0):
        sig = _textinfo_signature(page)
        if sig and sig != previous_sig:
            return True
        time.sleep(0.15)
    return False
def list_pli_side_nav_items(page):
    fr = _find_frame_with_selector(page,
        "xpath=//a[contains(@class,'GUIDE-sideNav')] | //div[contains(@class,'data-wrapper')]//a")
    if not fr:
        return []
    links = fr.locator("xpath=(//a[contains(@class,'GUIDE-sideNav')] | //div[contains(@class,'data-wrapper')]//a)")
    n = links.count()
    out = []
    for i in range(n):
        el = links.nth(i)
        t = clean(el.inner_text())
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
def read_analysis_summary_for_current_pli(page):
    fr, tbl = _find_latest_analysis_table_nearby(page)
    if not (fr and tbl and tbl.count()):
        return ""
    _expand_if_truncated(fr)
    wanted = [
        "Summary of Investigations",           # seen in your run
        "Analysis/Investigation Summary",
        "Analysis Summary",
        "Investigation Summary",
        "Analysis/Investigation conclusion",
        "Analysis/Investigation"
    ]
    row = None
    for t in wanted:
        cand = tbl.locator(
            "xpath=.//tr[td[starts-with(@id,'GUIDE-TextInfoTable-') and contains(@id,'-TextType') and normalize-space(.)=$t]]"
        ).filter(has_text=t).first
        if cand.count():
            row = cand
            break
    if not row or not row.count():
        for t in wanted:
            low = t.lower()
            cand = tbl.locator(
                "xpath=.//tr[td[starts-with(@id,'GUIDE-TextInfoTable-') and contains(@id,'-TextType') "
                f"and contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{low}')]]"
            ).first
            if cand.count():
                row = cand
                break
    if not row or not row.count():
        row = tbl.locator("xpath=.//tr[td[starts-with(@id,'GUIDE-TextInfoTable-') and contains(@id,'-Text')]]").first
        if not row.count():
            return ""
    td = row.locator("xpath=.//td[starts-with(@id,'GUIDE-TextInfoTable-') and contains(@id,'-Text')]").first
    if not td.count():
        return ""
    raw = td.evaluate("el => el.textContent || ''")
    text = re.sub(r'\r?\n\s*\r?\n+', '\n\n', raw)
    text = re.sub(r'[ \t\xa0]+', ' ', text)
    text = re.sub(r'\s*\n\s*', '\n', text).strip()
    return text
def collect_product_analysis(page, root_frame, known_products):
    had_click = click_left_nav_product_analysis(page)
    default_msg = "Information provided to Medtronic indicated that the complaint device was not available for evaluation."
    if not had_click:
        return { (p.get("id") or p.get("code") or "").strip(): default_msg
                 for p in known_products if (p.get("id") or p.get("code")) }
    items = list_pli_side_nav_items(page)
    from collections import defaultdict
    by_code = defaultdict(list)
    for it in items:
        by_code[it["code"]].append(it)
    results = {}
    for p in known_products:
        pid  = (p.get("id") or "").strip()
        code = (p.get("code") or extract_product_code(p.get("desc",""))).strip().upper()
        key  = code or pid
        if not key:
            continue
        candidates = []
        if code and code in by_code:
            candidates = by_code[code]
        else:
            tokens = [code] + re.findall(r'[A-Z0-9_]+', p.get("desc","").upper())
            for it in items:
                if any(tok and tok in it["text"].upper() for tok in tokens):
                    candidates.append(it)
        if not candidates:
            results[pid or code] = default_msg
            continue
        summaries = []
        for link in candidates:
            prev_sig = _textinfo_signature(page)           # snapshot current content
            if not robust_click(link["el"], link["frame"]):
                continue
            changed = False
            try:
                link["frame"].wait_for_selector(
                    "xpath=//div[contains(@class,'th-clr-cnt-bottom')]//td[starts-with(@id,'GUIDE-TextInfoTable-') "
                    "and contains(@id,'-TextType') and "
                    "(contains(.,'Summary of Investigations') or contains(.,'Analysis') or contains(.,'Investigation'))]",
                    timeout=6000
                )
                changed = True
            except Exception:
                pass
            if not changed:
                changed = wait_for_textinfo_change(page, prev_sig, timeout=8000)
            if not changed:
                link["frame"].wait_for_timeout(400)
            txt = read_analysis_summary_for_current_pli(page)
            if txt:
                summaries.append(txt)
                results[pid or code] = "\n\n".join(summaries) if summaries else default_msg
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
        page.goto(cfg['crm_url'], wait_until="load")
        sso_wait = cfg.get('sso_pause_seconds', 0)
        if sso_wait > 0:
            print(f"Waiting up to {sso_wait}s for SSO/MFA...")
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
                    target, target_ctx, used_sel = wait_find_in_any_frame(page, all_selectors, timeout_ms=s.get('pre_wait_timeout', 30000))
                    print(f"[Search] Found input via selector: {used_sel} in frame url={getattr(target_ctx, 'url', '')} name={getattr(target_ctx, 'name', '')}")
                    if s.get('clear', True):
                        try:
                            target.fill("")
                        except Exception:
                            pass
                    target.click()
                    frame.wait_for_timeout(100)
                    try:
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
                                btn.click()
                                submitted = True
                        except Exception:
                            pass
                    if not submitted:
                        key = s.get('press_key', 'Enter')
                        try: target.press(key)
                        except Exception: pass
                    if s.get('wait_for'):
                        target_ctx.wait_for_selector(s['wait_for'], timeout=s.get('wait_timeout', 60000))
                    else:
                        target_ctx.wait_for_load_state("networkidle")
                        target_ctx.wait_for_timeout(s.get('post_wait_ms', 3000))
                except Exception as e:
                    print(f"[Search] Failed to drive search: {e}")
                    dump_frames_debug(page, basename='debug')
                    try:
                        page.screenshot(path="debug_search_failure.png", full_page=True)
                        print("Saved debug_search_failure.png")
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
            if click_partners_tab(page, frame):
                pframe = find_partners_frame(page)
                if not pframe:
                    debug_frames_for_partners(page)
                    print("[Partners] Could not locate the partners frame.")
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
                    
            else:
                print("[Partners] Could not open Partners tab; leaving ir_* fields from label map/fallbacks.")
        except Exception as e:
            print(f"[Partners] Error scraping Partners tab: {e}")
        event_date_text = get_event_date(page)
        if event_date_text:
            values['event_date'] = event_date_text
        ext = read_external_refs(page, frame)
        if ext.get("rb_reference"):
            values["rb_reference"] = ext["rb_reference"]
        if ext.get("report_number"):
            values["report_number"] = ext["report_number"]
        products = read_all_products(page, frame)
        analysis_by_pid = collect_product_analysis(page, frame, products)
        prev_sig = _textinfo_signature(page)
        desc = read_event_description(page, frame)   # this opens Text Info tab
        if not desc:                                 # if first try races, wait and try again once
            wait_for_textinfo_change(page, prev_sig, timeout=8000)
            desc = read_event_description(page, frame)
        if desc:
            values["event_description"] = desc
        for idx, p in enumerate(products[:3], start=1):
            pid = (p.get("id") or "").strip()
            values[f"product_id_{idx}"]    = p.get("id","")
            values[f"product_desc_{idx}"]  = p.get("desc","")
            values[f"product_sn_{idx}"]    = p.get("sn","")
            values[f"product_lot_{idx}"]   = p.get("lot","")
            values[f"serial_or_lot_{idx}"] = " / ".join([s for s in [p.get('sn',''), p.get('lot','')] if s])
        lines = []
        for p in products:
            code = (p.get("code") or extract_product_code(p.get("desc",""))).upper()
            pid  = (p.get("id") or "").strip()
            summary = analysis_by_pid.get(pid) or analysis_by_pid.get(code, "")
            if summary:
                lines.append(f"{pid or code} — {summary}")
        values["analysis_results"] = "\n\n".join(lines) if lines else values.get("analysis_results","")
        if len(products) > 3:
            extras = [f"{p['id']} — {p['desc']}" for p in products[3:]]
            values.setdefault("product_extras", "\n".join(extras))
        print("Collected fields:")
        print(json.dumps(values, indent=2))
        out_name = cfg.get('output_name_pattern', 'Customer_Letter_{complaint_id}.docx').format(**values)
        out_path = out_dir / out_name
        fill_docx(str(template_path), str(out_path), values)
        print(f"Generated: {out_path}")
        #browser.close()
if __name__ == "__main__":
    main()