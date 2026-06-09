"""Minimal xlsx writer using only stdlib zipfile + xml."""
import zipfile, io, re
from datetime import date, datetime

# ── colour helpers ──────────────────────────────────────────────────────────
def _argb(hex6):
    return "FF" + hex6.upper().lstrip("#")

# ── shared-strings store ────────────────────────────────────────────────────
class Workbook:
    def __init__(self):
        self.sheets = []          # list of (name, Sheet)
        self._strings = []
        self._str_idx = {}

    def add_sheet(self, name):
        s = Sheet(self)
        self.sheets.append((name, s))
        return s

    def _si(self, s):
        if s not in self._str_idx:
            self._str_idx[s] = len(self._strings)
            self._strings.append(s)
        return self._str_idx[s]

    def save(self, path):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", self._content_types())
            zf.writestr("_rels/.rels", self._rels())
            zf.writestr("xl/workbook.xml", self._workbook_xml())
            zf.writestr("xl/_rels/workbook.xml.rels", self._wb_rels())
            zf.writestr("xl/styles.xml", self._styles_xml())
            zf.writestr("xl/sharedStrings.xml", self._shared_strings_xml())
            for i, (name, sheet) in enumerate(self.sheets, 1):
                zf.writestr(f"xl/worksheets/sheet{i}.xml", sheet._xml())
        data = buf.getvalue()
        # integrity check
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            bad = z.testzip()
            if bad:
                raise RuntimeError(f"Corrupt zip entry: {bad}")
        with open(path, "wb") as f:
            f.write(data)
        return data          # also return bytes so caller can base64-encode

    # ── package xml ──────────────────────────────────────────────────────────
    def _content_types(self):
        parts = ""
        for i in range(1, len(self.sheets)+1):
            parts += f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            + parts +
            '</Types>'
        )

    def _rels(self):
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>'
        )

    def _workbook_xml(self):
        sheets = ""
        for i, (name, _) in enumerate(self.sheets, 1):
            safe = name.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
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
        for i in range(1, len(self.sheets)+1):
            rels += (f'<Relationship Id="rId{i}" '
                     f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                     f'Target="worksheets/sheet{i}.xml"/>')
        rels += (
            f'<Relationship Id="rId{len(self.sheets)+1}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            'Target="styles.xml"/>'
            f'<Relationship Id="rId{len(self.sheets)+2}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
            'Target="sharedStrings.xml"/>'
        )
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + rels +
            '</Relationships>'
        )

    def _shared_strings_xml(self):
        items = ""
        for s in self._strings:
            safe = (s.replace("&","&amp;").replace("<","&lt;")
                     .replace(">","&gt;").replace('"',"&quot;"))
            items += f"<si><t xml:space=\"preserve\">{safe}</t></si>"
        c = len(self._strings)
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{c}" uniqueCount="{c}">'
            + items + '</sst>'
        )

    def _styles_xml(self):
        # Fonts: 0=Arial10, 1=Arial10 bold white, 2=Arial10 hyperlink
        fonts = (
            '<font><sz val="10"/><name val="Arial"/></font>'
            '<font><b/><sz val="10"/><name val="Arial"/><color rgb="FFFFFFFF"/></font>'
            '<font><sz val="10"/><name val="Arial"/><color rgb="FF0563C1"/><u val="single"/></font>'
        )
        # Fills: 0=none,1=gray(required),2=header(1F4E79),3=tier1(C8E6C9),4=tier2(FFF9C4)
        fills = (
            '<fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill>'
            '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/></patternFill></fill>'
            '<fill><patternFill patternType="solid"><fgColor rgb="FFC8E6C9"/></patternFill></fill>'
            '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF9C4"/></patternFill></fill>'
        )
        border = '<border><left/><right/><top/><bottom/><diagonal/></border>'
        # Cell xfs (styles):
        # 0 = default Arial10
        # 1 = header (font1, fill2)
        # 2 = tier1  (font0, fill3)
        # 3 = tier2  (font0, fill4)
        # 4 = hyperlink (font2, fill none)
        # 5 = tier1 hyperlink (font2, fill3)
        # 6 = tier2 hyperlink (font2, fill4)
        # 7 = date fmt (font0, numFmtId=14)
        # 8 = header date (font1, fill2, numFmtId=14)
        # 9 = tier1 date (font0, fill3, numFmtId=14)
        # 10= tier2 date (font0, fill4, numFmtId=14)
        xfs = (
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'                           # 0
            '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" applyFont="1" applyFill="1"/>'# 1 header
            '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" applyFill="1"/>'              # 2 tier1
            '<xf numFmtId="0" fontId="0" fillId="4" borderId="0" applyFill="1"/>'              # 3 tier2
            '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="1"/>'              # 4 hyperlink
            '<xf numFmtId="0" fontId="2" fillId="3" borderId="0" applyFont="1" applyFill="1"/>'# 5 tier1 hyper
            '<xf numFmtId="0" fontId="2" fillId="4" borderId="0" applyFont="1" applyFill="1"/>'# 6 tier2 hyper
            '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>'     # 7 date
            '<xf numFmtId="14" fontId="1" fillId="2" borderId="0" applyFont="1" applyFill="1" applyNumberFormat="1"/>'# 8 hdr date
            '<xf numFmtId="14" fontId="0" fillId="3" borderId="0" applyFill="1" applyNumberFormat="1"/>'# 9 t1 date
            '<xf numFmtId="14" fontId="0" fillId="4" borderId="0" applyFill="1" applyNumberFormat="1"/>'# 10 t2 date
        )
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="3">' + fonts + '</fonts>'
            '<fills count="5">' + fills + '</fills>'
            '<borders count="1">' + border + '</borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="11">' + xfs + '</cellXfs>'
            '</styleSheet>'
        )


# ── Sheet ───────────────────────────────────────────────────────────────────
COL_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

def _col_letter(n):   # 1-based
    r = ""
    while n:
        n, rem = divmod(n-1, 26)
        r = COL_LETTERS[rem] + r
    return r

def _esc(s):
    return (str(s).replace("&","&amp;").replace("<","&lt;")
                  .replace(">","&gt;").replace('"',"&quot;"))

_DATE_EPOCH = date(1899, 12, 30)

def _date_serial(d):
    if isinstance(d, datetime):
        d = d.date()
    return (d - _DATE_EPOCH).days


class Sheet:
    def __init__(self, wb):
        self.wb = wb
        self._rows = {}          # row_num -> {col_num -> cell_dict}
        self._col_widths = {}    # col_num -> width
        self._row_heights = {}   # row_num -> height
        self._hyperlinks = {}    # "A1" -> url
        self._freeze_row = None
        self._merge_cells = []

    # public API ────────────────────────────────────────────────────────────
    def write(self, row, col, value, style=0):
        """row,col are 1-based."""
        self._rows.setdefault(row, {})[col] = {"v": value, "s": style}

    def write_hyperlink(self, row, col, display, url, style=4):
        self._rows.setdefault(row, {})[col] = {"v": display, "s": style, "hl": True}
        ref = f"{_col_letter(col)}{row}"
        self._hyperlinks[ref] = url

    def set_col_width(self, col, width):
        self._col_widths[col] = width

    def set_row_height(self, row, height):
        self._row_heights[row] = height

    def freeze_row(self, row):
        self._freeze_row = row

    # xml ───────────────────────────────────────────────────────────────────
    def _xml(self):
        cols_xml = self._cols_xml()
        rows_xml = self._rows_xml()
        freeze_xml = self._freeze_xml()
        hl_xml = self._hl_xml()
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheetData>' + cols_xml + rows_xml + '</sheetData>'
            + freeze_xml + hl_xml +
            '</worksheet>'
        )

    def _cols_xml(self):
        if not self._col_widths:
            return ""
        parts = []
        for col, w in sorted(self._col_widths.items()):
            parts.append(f'<col min="{col}" max="{col}" width="{w}" customWidth="1"/>')
        return "<cols>" + "".join(parts) + "</cols>"

    def _rows_xml(self):
        # Rebuild: cols inside sheetData before rows is wrong; fix structure:
        # <sheetData> contains <row> elements only
        # We already return cols separately — but we wrapped rows in sheetData.
        # Let's fix: cols go BEFORE sheetData.
        return ""  # placeholder — handled in _xml2

    def _freeze_xml(self):
        if not self._freeze_row:
            return ""
        return (
            '<sheetViews><sheetView workbookViewId="0">'
            f'<pane ySplit="{self._freeze_row}" topLeftCell="A{self._freeze_row+1}" '
            'activePane="bottomLeft" state="frozen"/>'
            '</sheetView></sheetViews>'
        )

    def _hl_xml(self):
        if not self._hyperlinks:
            return ""
        parts = []
        for i, (ref, url) in enumerate(self._hyperlinks.items(), 1):
            safe_url = _esc(url)
            parts.append(f'<hyperlink ref="{ref}" r:id="rId{i}"/>')
        # We need relationships too — store for later
        self._hl_list = list(self._hyperlinks.items())
        return "<hyperlinks>" + "".join(parts) + "</hyperlinks>"

    def _xml(self):
        """Full sheet XML with correct element ordering."""
        # cols
        cols_xml = ""
        if self._col_widths:
            parts = []
            for col, w in sorted(self._col_widths.items()):
                parts.append(f'<col min="{col}" max="{col}" width="{w:.1f}" customWidth="1"/>')
            cols_xml = "<cols>" + "".join(parts) + "</cols>"

        # sheetViews (freeze)
        sv_xml = ""
        if self._freeze_row:
            sv_xml = (
                '<sheetViews><sheetView workbookViewId="0">'
                f'<pane ySplit="{self._freeze_row}" topLeftCell="A{self._freeze_row+1}" '
                'activePane="bottomLeft" state="frozen"/>'
                '</sheetView></sheetViews>'
            )

        # rows
        rows_xml = ""
        for rn in sorted(self._rows):
            h = self._row_heights.get(rn, "")
            h_attr = f' ht="{h}" customHeight="1"' if h else ""
            cells_xml = ""
            for cn in sorted(self._rows[rn]):
                cell = self._rows[rn][cn]
                ref = f"{_col_letter(cn)}{rn}"
                v = cell["v"]
                s = cell["s"]
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

        # hyperlinks
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
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            + sv_xml + cols_xml +
            '<sheetData>' + rows_xml + '</sheetData>'
            + hl_xml +
            '</worksheet>'
        )


# ── override save to also write sheet rels ──────────────────────────────────
_orig_save = Workbook.save

def _save_with_rels(self, path):
    # build XML for each sheet first so _hl_list is populated
    sheet_xmls = [(name, sheet._xml()) for name, sheet in self.sheets]

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", self._content_types())
        zf.writestr("_rels/.rels", self._rels())
        zf.writestr("xl/workbook.xml", self._workbook_xml())
        zf.writestr("xl/_rels/workbook.xml.rels", self._wb_rels())
        zf.writestr("xl/styles.xml", self._styles_xml())
        zf.writestr("xl/sharedStrings.xml", self._shared_strings_xml())
        for i, (name, sxml) in enumerate(sheet_xmls, 1):
            zf.writestr(f"xl/worksheets/sheet{i}.xml", sxml)
            sheet = self.sheets[i-1][1]
            if hasattr(sheet, "_hl_list") and sheet._hl_list:
                rels = ""
                for j, (ref, url) in enumerate(sheet._hl_list, 1):
                    safe_url = url.replace("&", "&amp;")
                    rels += (f'<Relationship Id="rId{j}" '
                             f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                             f'Target="{safe_url}" TargetMode="External"/>')
                zf.writestr(
                    f"xl/worksheets/_rels/sheet{i}.xml.rels",
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    + rels + '</Relationships>'
                )
    data = buf.getvalue()
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        bad = z.testzip()
        if bad:
            raise RuntimeError(f"Corrupt zip entry: {bad}")
    with open(path, "wb") as f:
        f.write(data)
    return data

Workbook.save = _save_with_rels
