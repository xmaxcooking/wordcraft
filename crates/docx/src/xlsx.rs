//! Minimal `.xlsx` cell reading (SpreadsheetML, ECMA-376 Part 1 §18): just enough of a workbook —
//! sheet names, shared strings, a worksheet's cell grid — to bake a chart's cache from it. Never
//! writes a workbook; the file's own bytes are kept unchanged as a chart's embedded source
//! (see `read::chart_insert`).

use crate::DocxError;
use crate::package::Package;
use crate::xml::El;

/// Most rows and columns read from one worksheet (hostile files are bounded).
pub(crate) const MAX_ROWS: usize = 16_384;
pub(crate) const MAX_COLS: usize = 1024;
/// Most shared strings read from the workbook.
const MAX_SHARED_STRINGS: usize = 200_000;

/// One cell: its display text, and its numeric value when it has one.
#[derive(Clone, Debug, Default)]
pub(crate) struct Cell {
    pub text: String,
    pub number: Option<f64>,
}

/// A worksheet's used cells, 0-based and rectangular (short rows pad with empty cells).
#[derive(Clone, Debug, Default)]
pub(crate) struct Grid {
    rows: Vec<Vec<Cell>>,
}

impl Grid {
    pub fn get(&self, row: usize, col: usize) -> Option<&Cell> {
        self.rows.get(row)?.get(col)
    }
    /// The bounding box of every non-empty cell: `(last_row, last_col)`, 0-based. `None` when
    /// the sheet has no cells.
    pub fn bounds(&self) -> Option<(usize, usize)> {
        let mut found = false;
        let (mut last_row, mut last_col) = (0usize, 0usize);
        for (r, row) in self.rows.iter().enumerate() {
            for (c, cell) in row.iter().enumerate() {
                if !cell.text.is_empty() || cell.number.is_some() {
                    found = true;
                    last_row = last_row.max(r);
                    last_col = last_col.max(c);
                }
            }
        }
        found.then_some((last_row, last_col))
    }
}

/// Reads sheet `sheet` (by name, or the workbook's first sheet when `None`) of an `.xlsx` package
/// into a dense cell grid.
pub(crate) fn read_sheet(xlsx: &[u8], sheet: Option<&str>) -> Result<Grid, DocxError> {
    let pkg = Package::open(xlsx)?;
    let workbook_path = "xl/workbook.xml";
    let wb = pkg.xml(workbook_path)?.ok_or_else(|| DocxError::MissingPart(workbook_path.into()))?;
    let sheets = wb.child("ss:sheets").ok_or_else(|| DocxError::Xml("workbook has no sheets".into()))?;
    let entry = match sheet {
        Some(name) => sheets.children("ss:sheet").find(|s| s.attr("name") == Some(name)),
        None => sheets.children("ss:sheet").next(),
    };
    let entry = entry.ok_or_else(|| match sheet {
        Some(name) => DocxError::Xml(format!("no sheet named {name:?}")),
        None => DocxError::Xml("the workbook has no sheets".into()),
    })?;
    let rid = entry.attr("r:id").ok_or_else(|| DocxError::Xml("a sheet has no relationship id".into()))?;
    let rels = pkg.rels(workbook_path);
    let rel = rels.by_id(rid).ok_or_else(|| DocxError::Xml("a sheet's relationship is missing".into()))?;
    let target = rel.target.clone();
    let sheet_xml = pkg.xml(&target)?.ok_or(DocxError::MissingPart(target))?;
    let strings = shared_strings(&pkg)?;
    Ok(parse_sheet(&sheet_xml, &strings))
}

fn shared_strings(pkg: &Package) -> Result<Vec<String>, DocxError> {
    let Some(sst) = pkg.xml("xl/sharedStrings.xml")? else { return Ok(Vec::new()) };
    Ok(sst.children("ss:si").take(MAX_SHARED_STRINGS).map(si_text).collect())
}

/// The text of a shared-string item (`ss:si`): direct text, or its rich-text runs concatenated.
fn si_text(si: &El) -> String {
    if let Some(t) = si.child("ss:t") {
        return t.text();
    }
    let mut s = String::new();
    for r in si.children("ss:r") {
        if let Some(t) = r.child("ss:t") {
            s.push_str(&t.text());
        }
    }
    s
}

pub(crate) fn parse_sheet(sheet: &El, strings: &[String]) -> Grid {
    let Some(data) = sheet.child("ss:sheetData") else { return Grid::default() };
    let mut rows: Vec<Vec<Cell>> = Vec::new();
    for (seq, row) in data.children("ss:row").enumerate() {
        if seq >= MAX_ROWS {
            break;
        }
        let r = row.attr("r").and_then(|v| v.parse::<usize>().ok()).unwrap_or(seq + 1).saturating_sub(1).min(MAX_ROWS - 1);
        if rows.len() <= r {
            rows.resize(r + 1, Vec::new());
        }
        for c in row.children("ss:c") {
            let Some(Some((col, _))) = c.attr("r").map(parse_ref) else { continue };
            if col >= MAX_COLS {
                continue;
            }
            let cell = parse_cell(c, strings);
            let cells = &mut rows[r];
            if cells.len() <= col {
                cells.resize(col + 1, Cell::default());
            }
            cells[col] = cell;
        }
    }
    Grid { rows }
}

fn parse_cell(c: &El, strings: &[String]) -> Cell {
    match c.attr("t").unwrap_or("n") {
        "s" => {
            let i = c.child("ss:v").and_then(|v| v.text().trim().parse::<usize>().ok());
            Cell { text: i.and_then(|i| strings.get(i)).cloned().unwrap_or_default(), number: None }
        }
        "str" | "e" => Cell { text: c.child("ss:v").map(|v| v.text()).unwrap_or_default(), number: None },
        "inlineStr" => Cell { text: c.child("ss:is").map(si_text).unwrap_or_default(), number: None },
        "b" => {
            let on = c.child("ss:v").map(|v| v.text()).is_some_and(|v| v.trim() == "1");
            Cell { text: if on { "TRUE".into() } else { "FALSE".into() }, number: None }
        }
        _ => {
            let n = c.child("ss:v").map(|v| v.text()).and_then(|v| v.trim().parse::<f64>().ok()).filter(|n| n.is_finite());
            Cell { text: n.map(|n| format!("{n}")).unwrap_or_default(), number: n }
        }
    }
}

/// `"A1"` → `(col, row)`, 0-based (`A` = 0, `Z` = 25, `AA` = 26…). `None` when it isn't one.
pub(crate) fn parse_ref(s: &str) -> Option<(usize, usize)> {
    let end = s.find(|c: char| !c.is_ascii_alphabetic())?;
    let (letters, digits) = (s.get(..end)?, s.get(end..)?);
    if letters.is_empty() || letters.len() > 3 || digits.is_empty() || !digits.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let mut col: usize = 0;
    for c in letters.chars() {
        col = col.checked_mul(26)?.checked_add(c.to_ascii_uppercase() as usize - 'A' as usize + 1)?;
    }
    let row: usize = digits.parse().ok()?;
    Some((col.checked_sub(1)?, row.checked_sub(1)?))
}

/// The column letters of 0-based column `col` (`0` → `"A"`, `26` → `"AA"`).
pub(crate) fn col_letters(col: usize) -> String {
    let mut n = col as u64 + 1;
    let mut s = Vec::new();
    while n > 0 {
        let r = (n - 1) % 26;
        s.push(b'A' + r as u8);
        n = (n - 1) / 26;
    }
    s.reverse();
    String::from_utf8(s).unwrap_or_default()
}

/// Parses an `"A1:C4"` style range (a single cell is also a range) into 0-based, inclusive
/// `(col0, row0, col1, row1)` bounds, low corner first.
pub(crate) fn parse_range(s: &str) -> Option<(usize, usize, usize, usize)> {
    let s = s.rsplit_once('!').map_or(s, |(_, rest)| rest);
    let (a, b) = s.split_once(':').unwrap_or((s, s));
    let (c0, r0) = parse_ref(a.trim())?;
    let (c1, r1) = parse_ref(b.trim())?;
    Some((c0.min(c1), r0.min(r1), c0.max(c1), r0.max(r1)))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cell_refs_round_trip() {
        for (s, col, row) in [("A1", 0, 0), ("B1", 1, 0), ("Z1", 25, 0), ("AA1", 26, 0), ("A2", 0, 1)] {
            assert_eq!(parse_ref(s), Some((col, row)), "{s}");
            assert_eq!(col_letters(col), s.trim_end_matches(char::is_numeric), "{s}");
        }
        assert_eq!(parse_ref(""), None);
        assert_eq!(parse_ref("1A"), None);
        assert_eq!(parse_ref("A"), None);
        assert_eq!(parse_ref("AAAA1"), None);
    }

    #[test]
    fn ranges_parse_low_corner_first() {
        assert_eq!(parse_range("B2:A1"), Some((0, 0, 1, 1)));
        assert_eq!(parse_range("Sheet1!A1:C4"), Some((0, 0, 2, 3)));
        assert_eq!(parse_range("A1"), Some((0, 0, 0, 0)));
        assert_eq!(parse_range("not a range"), None);
    }
}
