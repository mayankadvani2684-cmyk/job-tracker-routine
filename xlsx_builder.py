"""Minimal xlsx writer using only stdlib zipfile + xml."""
import zipfile, io
from datetime import date, datetime

# ── helpers ─────────────────────────────────────────────────────────────────

COL_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

def _col_letter(n):   # 1-based
    r = ""
    while n:
        n, rem = divmod(n - 1, 26)
        r = COL_LETTERS[rem] + r
    return r

def _esc(s):
    return (str(s)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))

_DATE_EPOCH = date(1899, 12, 30)

def _date_serial(d):
    if isinstance(d, datetime):
        d = d.date()
    return (d - _DATE_EPOCH).days


# ── Sheet ────────────────────────────────────────────────────────────────────

class Sheet:
    def __init__(self, wb):
        self.wb          = wb
        self._rows       = {}   # row_num -> {col_num -> cell_dict}
        self._col_widths = {}   # col_num -> width
        self._row_heights= {}   # row_num -> height
        self._hyperlinks = {}   # "A1"    -> url
        self._freeze_row = None

    # ── public API ───────────────────────────────────────────────────────────

    def write(self, row, col, value, style=0):
        """row, col are 1-based."""
        self._rows.setdefault(row, {})[col] = {"v": value, "s": style}

    def write_hyperlink(self, row, col, display, url, style=4):
        self._rows.setdefault(row, {})[col] = {"v": display, "s": style}
        ref = f"{_col_letter(col)}{row}"
        self._hyperlinks[ref] = url

    def set_col_width(self, col, width):
        self._col_widths[col] = width

    def set_row_height(self, row, height):
        self._row_heights[row] = height

    def freeze_row(self, row):
        self._freeze_row = row

    # ── xml generation ───────────────────────────────────────────────────────

    def _xml(self):
        """Build complete worksheet XML. Populates self._hl_list as a side effect."""

        # sheetViews (freeze pane)
        sv_xml = ""
        if self._freeze_row:
            sv_xml = (
                '<sheetViews><sheetView workbookViewId="0">'
                f'<pane ySplit="{self._freeze_row}" '
                f'topLeftCell="A{self._freeze_row + 1}" '
                'activePane="bottomLeft" state="frozen"/>'
                '</sheetView></sheetViews>'
            )

        # cols
        cols_xml = ""
        if self._col_widths:
            parts = []
            for col, w in sorted(self._col_widths.items()):
                parts.append(f'<col min="{col}" max="{col}" width="{w:.1f}" customWidth="1"/>')
            cols_xml = "<cols>" + "".join(parts) + "</cols>"

        # rows
        rows_xml = ""
        for rn in sorted(self._rows):
            h = self._row_heights.get(rn, "")
            h_attr = f' ht="{h}" customHeight="1"' if h else ""
            cells_xml = ""
            for cn in sorted(self._rows[rn]):
                cell = self._rows[rn][cn]
                ref  = f"{_col_letter(cn)}{rn}"
                v, s = cell["v"], cell["s"]
                if v is None or v == "":
                    cells_xml += f'<c r="{ref}" s="{s}"/>'
                elif isinstance(v, (date, datetime)):
                    serial = _date_serial(v)
                    cells_xml += f'<c r="{ref}" t="n" s="{s}"><v>{serial}</v></c>'
                elif isinstance(v, (int, float)):
                    cells_xml += f'<c r="{ref}" t="n" s="{s}"><v>{v}</v></c>'
                else:
                    si = self.wb._si(str(v))
                    cells_xml += f'<c r="{ref}" t="s" s="{s}"><v>{si}</v></c>'
            rows_xml += f'<row r="{rn}"{h_attr}>{cells_xml}</row>'

        # hyperlinks — populate _hl_list for save() to use when writing rels
        hl_xml = ""
        self._hl_list = []
        if self._hyperlinks:
            parts = []
            for i, (ref, url) in enumerate(self._hyperlinks.items(), 1):
                parts.append(f'<hyperlink ref="{ref}" r:id="rId{i}"/>')
                self._hl_list.append((ref, url))
            hl_xml = "<hyperlinks>" + "".join(parts) + "</hyperlinks>"

        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet '
            'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            + sv_xml + cols_xml
            + '<sheetData>' + rows_xml + '</sheetData>'
            + hl_xml
            + '</worksheet>'
        )


# ── Workbook ─────────────────────────────────────────────────────────────────

class Workbook:
    def __init__(self):
        self.sheets    = []   # list of (name, Sheet)
        self._strings  = []
        self._str_idx  = {}

    def add_sheet(self, name):
        s = Sheet(self)
        self.sheets.append((name, s))
        return s

    def _si(self, s):
        """Return shared-string index for s, inserting if necessary."""
        if s not in self._str_idx:
            self._str_idx[s] = len(self._strings)
            self._strings.append(s)
        return self._str_idx[s]

    # ── save ─────────────────────────────────────────────────────────────────

    def save(self, path):
        """
        Build the xlsx zip in memory, verify integrity, write to path.
        Sheet rels (hyperlinks) are written in the same pass as sheet XML —
        no monkey-patching, no second pass.
        Returns the raw bytes so callers can base64-encode if needed.
        """
        # Generate sheet XML first so _hl_list is populated on each Sheet
        sheet_xmls = [(name, sheet._xml()) for name, sheet in self.sheets]

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml",      self._content_types())
            zf.writestr("_rels/.rels",               self._rels())
            zf.writestr("xl/workbook.xml",           self._workbook_xml())
            zf.writestr("xl/_rels/workbook.xml.rels",self._wb_rels())
            zf.writestr("xl/styles.xml",             self._styles_xml())
            zf.writestr("xl/sharedStrings.xml",      self._shared_strings_xml())

            for i, (name, sxml) in enumerate(sheet_xmls, 1):
                zf.writestr(f"xl/worksheets/sheet{i}.xml", sxml)

                # Write sheet rels only if this sheet has hyperlinks
                sheet = self.sheets[i - 1][1]
                hl_list = getattr(sheet, "_hl_list", [])
                if hl_list:
                    rels_parts = []
                    for j, (ref, url) in enumerate(hl_list, 1):
                        safe_url = url.replace("&", "&amp;")
                        rels_parts.append(
                            f'<Relationship Id="rId{j}" '
                            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                            f'Target="{safe_url}" TargetMode="External"/>'
                        )
                    zf.writestr(
                        f"xl/worksheets/_rels/sheet{i}.xml.rels",
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                        + "".join(rels_parts)
                        + '</Relationships>'
                    )

        data = buf.getvalue()

        # Integrity check — fail loudly before touching disk
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            bad = z.testzip()
            if bad:
                raise RuntimeError(f"Corrupt zip entry after build: {bad}")

        with open(path, "wb") as f:
            f.write(data)

        return data   # also return bytes so caller can base64-encode for Drive upload

    # ── package XML helpers ───────────────────────────────────────────────────

    def _content_types(self):
        parts = ""
        for i in range(1, len(self.sheets) + 1):
            parts += (
                f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                f'ContentType="application/vnd.openxmlformats-officedocument'
                f'.spreadsheetml.worksheet+xml"/>'
            )
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            + parts
            + '</Types>'
        )

    def _rels(self):
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/>'
            '</Relationships>'
        )

    def _workbook_xml(self):
        sheets = ""
        for i, (name, _) in enumerate(self.sheets, 1):
            safe = name.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            sheets += f'<sheet name="{safe}" sheetId="{i}" r:id="rId{i}"/>'
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets>' + sheets + '</sheets>'
            '</workbook>'
        )

    def _wb_rels(self):
        rels = ""
        for i in range(1, len(self.sheets) + 1):
            rels += (
                f'<Relationship Id="rId{i}" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                f'Target="worksheets/sheet{i}.xml"/>'
            )
        n = len(self.sheets)
        rels += (
            f'<Relationship Id="rId{n + 1}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            'Target="styles.xml"/>'
            f'<Relationship Id="rId{n + 2}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
            'Target="sharedStrings.xml"/>'
        )
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + rels
            + '</Relationships>'
        )

    def _shared_strings_xml(self):
        items = ""
        for s in self._strings:
            safe = (_esc(s))
            items += f'<si><t xml:space="preserve">{safe}</t></si>'
        c = len(self._strings)
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'count="{c}" uniqueCount="{c}">'
            + items
            + '</sst>'
        )

    def _styles_xml(self):
        # Fonts
        # 0 = Arial 10 normal
        # 1 = Arial 10 bold white  (header)
        # 2 = Arial 10 hyperlink blue underline
        fonts = (
            '<font><sz val="10"/><name val="Arial"/></font>'
            '<font><b/><sz val="10"/><name val="Arial"/><color rgb="FFFFFFFF"/></font>'
            '<font><sz val="10"/><name val="Arial"/><color rgb="FF0563C1"/><u val="single"/></font>'
        )

        # Fills (indices 0-1 are required by spec)
        # 0 = none, 1 = gray125, 2 = header navy, 3 = tier1 green, 4 = tier2 yellow
        fills = (
            '<fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill>'
            '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/></patternFill></fill>'
            '<fill><patternFill patternType="solid"><fgColor rgb="FFC8E6C9"/></patternFill></fill>'
            '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF9C4"/></patternFill></fill>'
        )

        border = '<border><left/><right/><top/><bottom/><diagonal/></border>'

        # Cell xf styles (index = style number passed to write/write_hyperlink)
        # 0  = default
        # 1  = header      (font1 bold-white, fill2 navy)
        # 2  = tier1 data  (font0, fill3 green)
        # 3  = tier2 data  (font0, fill4 yellow)
        # 4  = hyperlink plain        (font2 blue, fill0)
        # 5  = tier1 hyperlink        (font2 blue, fill3 green)
        # 6  = tier2 hyperlink        (font2 blue, fill4 yellow)
        # 7  = date plain             (font0, fill0, numFmt date)
        # 8  = date header            (font1, fill2, numFmt date)
        # 9  = tier1 date             (font0, fill3, numFmt date)
        # 10 = tier2 date             (font0, fill4, numFmt date)
        xfs = (
            '<xf numFmtId="0"  fontId="0" fillId="0" borderId="0"/>'
            '<xf numFmtId="0"  fontId="1" fillId="2" borderId="0" applyFont="1" applyFill="1"/>'
            '<xf numFmtId="0"  fontId="0" fillId="3" borderId="0" applyFill="1"/>'
            '<xf numFmtId="0"  fontId="0" fillId="4" borderId="0" applyFill="1"/>'
            '<xf numFmtId="0"  fontId="2" fillId="0" borderId="0" applyFont="1"/>'
            '<xf numFmtId="0"  fontId="2" fillId="3" borderId="0" applyFont="1" applyFill="1"/>'
            '<xf numFmtId="0"  fontId="2" fillId="4" borderId="0" applyFont="1" applyFill="1"/>'
            '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>'
            '<xf numFmtId="14" fontId="1" fillId="2" borderId="0" applyFont="1" applyFill="1" applyNumberFormat="1"/>'
            '<xf numFmtId="14" fontId="0" fillId="3" borderId="0" applyFill="1" applyNumberFormat="1"/>'
            '<xf numFmtId="14" fontId="0" fillId="4" borderId="0" applyFill="1" applyNumberFormat="1"/>'
        )

        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="3">'   + fonts  + '</fonts>'
            '<fills count="5">'   + fills  + '</fills>'
            '<borders count="1">' + border + '</borders>'
            '<cellStyleXfs count="1">'
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
            '</cellStyleXfs>'
            '<cellXfs count="11">' + xfs + '</cellXfs>'
            '</styleSheet>'
        )
