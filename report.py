"""
PR Lookup — enriches a Purchase Requisition file against Set_Master_Combined.

Input  : PR file (.xlsx) with exactly these 10 columns:
           Purchase Requsition Number | Item Id | CW Qty | Store | Priority |
           Site | Warehouse | Location | Set Id | Merchandise Expected Date

SMC source (auto-detected, in priority order):
  1. AZURE_SMC_URL env var  → download parquet from Azure Blob
  2. Local  Set_Master_Combined.parquet  (same folder as this script)
  3. Local  Set_Master_Combined.xlsx     (same folder as this script)

Output : single Excel sheet —
  PR cols (10)  +  SMC cols (15)  +  Duplicate Flag col

SMC cols carried through (15 total):
  Set Code | Set Active Status | Line Number | Child Code |
  Child Active Set Membership | Child Inactive Set Membership | Type |
  Child Active Status | Inactive Reason (SearchName) |
  Alternate Code | Alternate Code Status | Alt Code Set Membership |
  Item Stage | Item Model Group | Ledger Dimension
"""

import sys, os, io, time, tempfile
import openpyxl
import pyarrow.parquet as pq
from openpyxl.styles      import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils       import get_column_letter
from openpyxl.cell.text   import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText
from python_calamine      import CalamineWorkbook
from collections          import defaultdict, Counter


# ── constants ─────────────────────────────────────────────────────────────────

AZURE_SMC_URL_DEFAULT = (
    "https://njlprodimages.blob.core.windows.net/protopisfolder/Set_Master_Combined.parquet"
)

REQUIRED_PR_COLS = [
    'Purchase Requsition Number', 'Item Id', 'CW Qty', 'Store',
    'Priority', 'Site', 'Warehouse', 'Location', 'Set Id',
    'Merchandise Expected Date',
]

# Number of SMC columns carried into the output (12 original + 3 new RP cols)
SMC_COL_COUNT = 15


# ── SMC loader (parquet-first, xlsx fallback) ─────────────────────────────────

def _smc_from_table(table):
    """
    Convert a PyArrow Table into (smc_header, by_set, by_child).

    Works column-by-column via to_pydict() so we never hold a full
    pandas DataFrame in memory — critical for Render's 512 MB free tier.

    Parquet schema (15 cols, 0-based):
      0  set_code                            1  set_active         2  linenum
      3  itemid (Child Code)                 4  child_active_set_membership
      5  child_inactive_set_membership       6  row_type (Type)    7  child_status
      8  child_searchname                    9  alt_code          10  alt_status
      11 alt_set_membership                 12  pwc_itemstage     13  itemmodelgroupid
      14 defaultledgerdimensiondisplayvalue
    """
    keep       = SMC_COL_COUNT
    all_cols   = table.column_names
    smc_header = all_cols[:keep]

    # Pull only the columns we need as plain Python lists — one allocation,
    # then the table itself can be GC'd.
    col_data = []
    for i in range(keep):
        col = table.column(i) if i < len(all_cols) else None
        if col is not None:
            col_data.append([
                '' if v is None else str(v).strip()
                for v in col.to_pylist()
            ])
        else:
            col_data.append([''] * table.num_rows)

    del table  # free the Arrow memory immediately

    by_set   = defaultdict(list)
    by_child = defaultdict(list)

    for row_idx in range(len(col_data[0])):
        r          = [col_data[c][row_idx] for c in range(keep)]
        set_code   = r[0]
        child_code = r[3]
        if set_code:
            by_set[set_code].append(r)
        if child_code:
            by_child[child_code].append(r)

    return smc_header, by_set, by_child


def _smc_from_df(df):
    """Thin shim for the xlsx fallback path (df is already built, keep it small)."""
    import pyarrow as pa
    table = pa.Table.from_pandas(df, preserve_index=False)
    return _smc_from_table(table)


def load_smc(path_or_url=None):
    """
    Load Set_Master_Combined from the best available source.

    Priority:
      1. Explicit path_or_url argument (if given)
      2. AZURE_SMC_URL  env var
      3. AZURE_SMC_URL_DEFAULT constant
      4. Local .parquet beside this script
      5. Local .xlsx beside this script
    """
    script_dir    = os.path.dirname(os.path.abspath(__file__))
    local_parquet = os.path.join(script_dir, "Set_Master_Combined.parquet")
    local_xlsx    = os.path.join(script_dir, "Set_Master_Combined.xlsx")

    # ── caller passed a local parquet path ───────────────────────────────────
    if path_or_url and os.path.isfile(path_or_url) and path_or_url.endswith(".parquet"):
        print(f"  Loading SMC from local parquet: {path_or_url}")
        table = pq.read_table(path_or_url)
        return _smc_from_table(table)

    # ── try Azure / env-var URL ───────────────────────────────────────────────
    azure_url = (
        path_or_url
        if (path_or_url and not os.path.isfile(path_or_url))
        else os.environ.get("AZURE_SMC_URL", AZURE_SMC_URL_DEFAULT)
    )

    try:
        import urllib.request
        print(f"  Fetching SMC from Azure: {azure_url}")
        with urllib.request.urlopen(azure_url, timeout=60) as resp:
            data = resp.read()
        # Read directly into Arrow — never builds a pandas DataFrame
        table = pq.read_table(io.BytesIO(data))
        nrows = table.num_rows
        del data   # free the raw bytes immediately before building index dicts
        print(f"  SMC loaded from Azure ({nrows:,} rows, {table.num_columns} cols)")
        return _smc_from_table(table)
    except Exception as e:
        print(f"  Azure fetch failed ({e}), trying local files…")

    # ── try local parquet ─────────────────────────────────────────────────────
    if os.path.isfile(local_parquet):
        print(f"  Loading SMC from local parquet: {local_parquet}")
        table = pq.read_table(local_parquet)
        return _smc_from_table(table)

    # ── fallback: local xlsx ──────────────────────────────────────────────────
    if path_or_url and os.path.isfile(path_or_url) and path_or_url.endswith(".xlsx"):
        src = path_or_url
    elif os.path.isfile(local_xlsx):
        src = local_xlsx
    else:
        raise FileNotFoundError(
            "Set_Master_Combined not found. "
            "Provide AZURE_SMC_URL env var or place .parquet/.xlsx beside report.py"
        )

    print(f"  Loading SMC from local xlsx: {src}")
    wb    = CalamineWorkbook.from_path(src)
    sheet = wb.get_sheet_by_index(0)
    rows  = sheet.to_python(skip_empty_area=False)
    if not rows:
        raise ValueError("Set_Master_Combined.xlsx is empty")

    raw_header = [str(h).strip() for h in rows[0]]
    keep = SMC_COL_COUNT
    df   = pd.DataFrame(
        [
            [str(v).strip() if v is not None else '' for v in (list(r) + [''] * (keep - len(r)))[:keep]]
            for r in rows[1:]
        ],
        columns=(raw_header + [''] * max(0, keep - len(raw_header)))[:keep],
    )
    return _smc_from_df(df)


# ── PR loader ─────────────────────────────────────────────────────────────────

def load_pr(path):
    wb    = CalamineWorkbook.from_path(path)
    sheet = wb.get_sheet_by_index(0)
    rows  = sheet.to_python(skip_empty_area=False)
    if not rows:
        return [], []
    header = [str(h).strip() for h in rows[0]]
    return header, rows[1:]


def validate_pr_headers(header):
    missing = [c for c in REQUIRED_PR_COLS if c not in header]
    if missing:
        raise ValueError(
            f"Input file is missing required columns: {', '.join(missing)}.\n"
            f"Expected: {', '.join(REQUIRED_PR_COLS)}"
        )
    return True


# ── style helpers ─────────────────────────────────────────────────────────────

def make_border(color='CCCCCC'):
    s = Side(style='thin', color=color)
    return Border(left=s, right=s, top=s, bottom=s)


# ── main processing ───────────────────────────────────────────────────────────

def run_lookup(pr_path, smc_source=None, output_path=None):
    """
    pr_path     : path to uploaded PR xlsx
    smc_source  : parquet URL / local path / None (auto-detect)
    output_path : where to write the result xlsx.
                  If None, writes to a temp file and returns its path.
    """
    t0 = time.time()
    print("Loading files…")

    pr_header, pr_rows = load_pr(pr_path)
    validate_pr_headers(pr_header)
    print(f"  PR rows: {len(pr_rows)}")

    smc_header, by_set, by_child = load_smc(smc_source)
    print(f"  SMC sets indexed: {len(by_set)}  |  children indexed: {len(by_child)}")
    print(f"  SMC columns     : {len(smc_header)}  ({', '.join(smc_header[-3:])})") # show last 3 to confirm new cols

    # ── column indices in PR ──────────────────────────────────────────────────
    def pr_ci(name):
        return pr_header.index(name) if name in pr_header else None

    ITEM_CI = pr_ci('Item Id')
    SET_CI  = pr_ci('Set Id')

    # ── duplicate analysis ────────────────────────────────────────────────────
    set_id_list  = [str(r[SET_CI]).strip()  for r in pr_rows if SET_CI  is not None and str(r[SET_CI]).strip()]
    item_id_list = [str(r[ITEM_CI]).strip() for r in pr_rows if ITEM_CI is not None and str(r[ITEM_CI]).strip()]

    dup_sets         = {k for k, v in Counter(set_id_list).items()  if v > 1}
    dup_items        = {k for k, v in Counter(item_id_list).items() if v > 1}
    set_id_set       = set(set_id_list)
    item_id_set      = set(item_id_list)
    cross_overlap    = set_id_set & item_id_set
    smc_child_set    = set(by_child.keys())
    set_also_child   = set_id_set  & smc_child_set
    item_also_set    = item_id_set & set(by_set.keys())

    print(f"  Duplicate Set Ids   : {len(dup_sets)}")
    print(f"  Duplicate Item Ids  : {len(dup_items)}")
    print(f"  Cross-overlap       : {len(cross_overlap)}")

    # ── build output rows ─────────────────────────────────────────────────────
    EMPTY_SMC = [''] * SMC_COL_COUNT
    out_rows  = []

    for pr_row in pr_rows:
        item_id = str(pr_row[ITEM_CI]).strip() if ITEM_CI is not None else ''
        set_id  = str(pr_row[SET_CI]).strip()  if SET_CI  is not None else ''

        flags = []
        if set_id and set_id in dup_sets:
            flags.append(f'Duplicate Set Id in PR ({set_id})')
        if item_id and item_id in dup_items:
            flags.append(f'Duplicate Item Id in PR ({item_id})')
        if set_id and set_id in cross_overlap:
            flags.append(f'Set Id also appears as Item Id ({set_id})')
        if item_id and item_id in cross_overlap:
            flags.append(f'Item Id also appears as Set Id ({item_id})')
        if set_id and set_id in set_also_child:
            flags.append(f'Set Id is also a Child Code in SMC ({set_id})')
        if item_id and item_id in item_also_set:
            flags.append(f'Item Id is also a Set Code in SMC ({item_id})')
        if item_id and len(by_child.get(item_id, [])) > 1:
            flags.append(f'Item Id has {len(by_child[item_id])} rows in SMC')

        dup_flag = ' | '.join(flags) if flags else ''

        if set_id:
            smc_matches = by_set.get(set_id, [])
            if not smc_matches:
                note = (dup_flag + ' | NOT FOUND IN SMC') if dup_flag else 'NOT FOUND IN SMC'
                out_rows.append((list(pr_row), list(EMPTY_SMC), note))
            else:
                for smc_row in smc_matches:
                    out_rows.append((list(pr_row), smc_row, dup_flag))
        elif item_id:
            smc_matches = by_child.get(item_id, [])
            if not smc_matches:
                note = (dup_flag + ' | NOT FOUND IN SMC') if dup_flag else 'NOT FOUND IN SMC'
                out_rows.append((list(pr_row), list(EMPTY_SMC), note))
            else:
                # col 6 = Type; prefer Standalone row if multiple matches
                preferred = next((r for r in smc_matches if str(r[6]).strip() == 'Standalone'), smc_matches[0])
                if len(smc_matches) > 1:
                    extra = f'Item Id found in {len(smc_matches)} SMC rows (showing Standalone)'
                    dup_flag = (dup_flag + ' | ' + extra) if dup_flag else extra
                out_rows.append((list(pr_row), preferred, dup_flag))
        else:
            out_rows.append((list(pr_row), list(EMPTY_SMC), dup_flag))

    print(f"  Output rows: {len(out_rows)}")

    # ── Write Excel ───────────────────────────────────────────────────────────
    # Friendly display headers for the 15 SMC cols
    SMC_DISPLAY_HEADERS = [
        'Set Code', 'Set Active Status', 'Line Number', 'Child Code',
        'Child Active Set Membership', 'Child Inactive Set Membership', 'Type',
        'Child Active Status', 'Inactive Reason (SearchName)',
        'Alternate Code', 'Alternate Code Status', 'Alt Code Set Membership',
        'Item Stage', 'Item Model Group', 'Ledger Dimension',
    ]
    all_headers = pr_header + SMC_DISPLAY_HEADERS + ['Duplicate / Note']

    wb_out = openpyxl.Workbook()
    ws     = wb_out.active
    ws.title = "PR Lookup"

    bdr     = make_border()
    h_fill  = PatternFill("solid", fgColor="1F3864")
    h_font  = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    h_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for ci, h in enumerate(all_headers, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font, c.fill, c.alignment, c.border = h_font, h_fill, h_align, bdr
    ws.row_dimensions[1].height = 30

    active_f   = PatternFill("solid", fgColor="C6EFCE")
    inactive_f = PatternFill("solid", fgColor="FFC7CE")
    notfound_f = PatternFill("solid", fgColor="FFEB9C")
    warn_f     = PatternFill("solid", fgColor="FFF2CC")
    dup_f      = PatternFill("solid", fgColor="FCE4D6")
    setrow_f   = PatternFill("solid", fgColor="EBF3FB")
    noset_f    = PatternFill("solid", fgColor="F2DCDB")
    type_set_f = PatternFill("solid", fgColor="E2EFDA")
    type_sa_f  = PatternFill("solid", fgColor="EDEDED")
    rp_fill    = PatternFill("solid", fgColor="EAF0FB")   # soft blue for new RP cols

    b_font     = Font(name="Arial", size=9)
    b_align    = Alignment(vertical="center")
    wrap_align = Alignment(vertical="top", wrap_text=True)

    PR_LEN    = len(pr_header)
    SMC_START = PR_LEN + 1           # 1-based column index of first SMC col

    # SMC column positions (1-based absolute)
    SMC_SET_ACTIVE_CI   = SMC_START + 1    # Set Active Status
    SMC_CHILD_STATUS_CI = SMC_START + 7    # Child Active Status
    SMC_ALT_STATUS_CI   = SMC_START + 10   # Alternate Code Status
    SMC_STATUS_COLS     = {SMC_SET_ACTIVE_CI, SMC_CHILD_STATUS_CI, SMC_ALT_STATUS_CI}
    SMC_TYPE_CI         = SMC_START + 6    # Type
    SMC_MEMBER_COLS     = {SMC_START + 4, SMC_START + 5, SMC_START + 11}  # membership cols
    SMC_RP_EXTRA_COLS   = {SMC_START + 12, SMC_START + 13, SMC_START + 14}  # new RP cols
    DUP_CI              = PR_LEN + SMC_COL_COUNT + 1   # last col

    def sfmt(val):
        s = (val or '').lower()
        if s == 'active':    return active_f,   "276221"
        if s == 'not found': return notfound_f, "9C5700"
        if s in ('na', ''):  return None, None
        return inactive_f, "9C0006"

    for ri, (pr_raw, smc_raw, dup_flag) in enumerate(out_rows, start=2):
        has_dup   = bool(dup_flag)
        is_set_row = bool(str(pr_raw[SET_CI]).strip()) if SET_CI is not None else False

        # col 7 (0-based) = child_status
        smc_child_status = str(smc_raw[7]).strip() if len(smc_raw) > 7 and smc_raw[7] else ''
        inactive_child   = smc_child_status.lower() not in ('active', '')

        base_fill = dup_f if has_dup else (setrow_f if is_set_row else None)
        if inactive_child and not has_dup:
            base_fill = warn_f

        # dynamic row height based on membership cell line counts
        csm_active_val   = smc_raw[4]  if len(smc_raw) > 4  else None
        csm_inactive_val = smc_raw[5]  if len(smc_raw) > 5  else None
        asm_val          = smc_raw[11] if len(smc_raw) > 11 else None
        csm_lines = max(
            str(csm_active_val).count('\n') + 1   if csm_active_val   else 1,
            str(csm_inactive_val).count('\n') + 1 if csm_inactive_val else 1,
        )
        asm_lines = str(asm_val).count('\n') + 1 if asm_val else 1
        ws.row_dimensions[ri].height = max(16, min(max(csm_lines, asm_lines) * 14, 80))

        # ── PR cols ───────────────────────────────────────────────────────────
        for ci, val in enumerate(pr_raw, start=1):
            v = val if val not in (None, '') else None
            cell = ws.cell(row=ri, column=ci, value=v)
            cell.font, cell.alignment, cell.border = b_font, b_align, bdr
            if base_fill:
                cell.fill = base_fill

        # ── SMC cols (all 15) ─────────────────────────────────────────────────
        for smc_i, val in enumerate(smc_raw):
            ci = SMC_START + smc_i   # 1-based
            is_rich = isinstance(val, CellRichText)
            v = val if (is_rich or val not in (None, '')) else None
            cell = ws.cell(row=ri, column=ci, value=v)
            cell.border = bdr

            if ci in SMC_MEMBER_COLS:
                cell.alignment = wrap_align
                cell.font      = Font(name="Arial", size=9)
                if base_fill:
                    cell.fill = base_fill
                if val == 'NOT A CHILD IN ANY SET':
                    cell.fill = noset_f
                    cell.font = Font(name="Arial", size=9, color="9C0006")
            elif ci in SMC_RP_EXTRA_COLS:
                # New RP enrichment cols — distinct soft-blue styling
                cell.font      = Font(name="Arial", size=9, color="1F3864")
                cell.alignment = b_align
                cell.fill      = rp_fill
            else:
                cell.font      = b_font
                cell.alignment = b_align
                if base_fill:
                    cell.fill = base_fill

            if ci in SMC_STATUS_COLS:
                fill, fc = sfmt(str(val) if val else '')
                if fill:
                    cell.fill = fill
                    cell.font = Font(name="Arial", size=9, color=fc, bold=True)

            if ci == SMC_TYPE_CI:
                tv = str(val).strip() if val else ''
                cell.fill      = type_set_f if tv == 'Set' else type_sa_f
                cell.font      = Font(name="Arial", size=9,
                                      color="375623" if tv == 'Set' else "595959",
                                      bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")

        # ── Duplicate flag col ────────────────────────────────────────────────
        dup_cell = ws.cell(row=ri, column=DUP_CI, value=dup_flag if dup_flag else None)
        dup_cell.border    = bdr
        dup_cell.alignment = Alignment(vertical="center", wrap_text=True)
        if dup_flag:
            dup_cell.fill = dup_f
            dup_cell.font = Font(name="Arial", size=9, color="C00000", bold=True)
        else:
            dup_cell.font = b_font

    # ── column widths ─────────────────────────────────────────────────────────
    pr_widths  = [28, 22, 10, 12, 12, 12, 12, 12, 22, 20]
    smc_widths = [22, 18, 20, 22, 28, 28, 14, 20, 28, 25, 22, 28, 20, 22, 30]
    dup_width  = [40]
    for i, w in enumerate(pr_widths + smc_widths + dup_width, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(all_headers))}1"

    # ── resolve output path ───────────────────────────────────────────────────
    if output_path is None:
        tmp = tempfile.NamedTemporaryFile(
            suffix=".xlsx", prefix="PR_Lookup_", delete=False
        )
        output_path = tmp.name
        tmp.close()

    out_dir = os.path.dirname(os.path.abspath(output_path))
    os.makedirs(out_dir, exist_ok=True)
    wb_out.save(output_path)

    t1 = time.time()
    dup_rows    = sum(1 for _, _, f in out_rows if f)
    set_expRows = sum(1 for pr, _, _ in out_rows if SET_CI and str(pr[SET_CI]).strip())
    print(f"  Excel written: {t1-t0:.2f}s")
    print(f"\n✅  Done in {t1-t0:.2f}s  →  {output_path}")
    print(f"   {len(out_rows)} rows  |  {set_expRows} set-expanded  |  {dup_rows} with flags")

    return output_path


# ── CLI entry point ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) >= 3:
        pr_p   = sys.argv[1]
        smc_p  = sys.argv[2]
        out_p  = sys.argv[3] if len(sys.argv) > 3 else None
        run_lookup(pr_p, smc_p, out_p)
    else:
        print("=" * 58)
        print("  PR LOOKUP — Provide Input File Paths")
        print("=" * 58)

        def ask_path(label):
            while True:
                p = input(f"\n{label}\n> ").strip().strip('"').strip("'")
                if os.path.isfile(p):
                    return p
                print(f"  File not found: {p!r}  — please try again.")

        pr_p  = ask_path("(1/2) PR input file path (.xlsx):")
        smc_p = input("\n(2/2) SMC path or leave blank to use Azure/auto:\n> ").strip() or None
        base  = os.path.splitext(os.path.basename(pr_p))[0]
        out_p = os.path.join(os.path.dirname(os.path.abspath(pr_p)), f"{base}_Lookup.xlsx")
        print(f"\nOutput → {out_p}\n")
        run_lookup(pr_p, smc_p, out_p)
