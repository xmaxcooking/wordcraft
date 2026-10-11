//! Insert a chart backed by an embedded `.xlsx` workbook, the way Word's Insert › Chart does: the
//! cache ([`super::chart`]'s `c:numCache`/`c:strCache`) is baked from the user's workbook once,
//! and the workbook's own bytes are copied in unchanged as the chart's `c:externalData` part. No
//! `.xlsx` writer, chart type picker or data-editing dialog: that's a separate problem (see
//! `docs/gaps.md` #13). Saving writes the chart back exactly as any other kept embedded object
//! ([`crate::write::embed`]), since it is one as far as that writer knows.

use std::sync::Arc;

use wordcraft_doc::Document;
use wordcraft_doc::graphic::{Embedded, EmbeddedRef, Graphic, GraphicKind};

use crate::DocxError;
use crate::package::{EMBEDDED_PARTS, EmbeddedManifest, EmbeddedPart, Rel, rt};
use crate::xlsx::{self, Grid};
use crate::xml::{self, W};

const CHART_CT: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const XLSX_CT: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/// Most series, and most categories per series, a baked chart holds (a hostile or huge range is
/// truncated, not refused).
const MAX_SERIES: usize = 255;
const MAX_CATEGORIES: usize = 2000;
/// The range read when none is given: the sheet's used cells, capped.
const MAX_DEFAULT_ROWS: usize = 50;
const MAX_DEFAULT_COLS: usize = 20;
/// Attempts at a free part name before giving up on prettiness.
const MAX_NAME_ATTEMPTS: usize = 10_000;

/// Chart types this can bake. Word's other chart types (and the data-grid editor of #409) are
/// future work.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Kind {
    Bar,
    Line,
    Area,
    Pie,
    Doughnut,
}

impl Kind {
    fn parse(s: &str) -> Kind {
        match s.to_ascii_lowercase().as_str() {
            "line" => Kind::Line,
            "area" => Kind::Area,
            "pie" => Kind::Pie,
            "doughnut" | "donut" => Kind::Doughnut,
            _ => Kind::Bar,
        }
    }
}

/// Categories and series values read from a sheet range, in the layout Excel and Word assume for
/// a selection with more than one row and column: the first row (from the second column) names
/// the series, the first column (from the second row) labels the categories. A single row or
/// column is one series, labelled by position.
struct ChartData {
    categories: Vec<String>,
    series: Vec<(String, Vec<Option<f64>>)>,
}

/// Inserts a chart of `kind` (`"bar"`, `"line"`, `"area"`, `"pie"` or `"doughnut"`; anything else
/// is `"bar"`) baked from `sheet` (the workbook's first sheet when `None`) of `xlsx`, over `range`
/// (an `"A1:C4"` style reference, or the sheet's used cells when `None`), titled `title`. `w`/`h`
/// are the chart's size in points. Returns the graphic to place at an inline object; `doc` gains
/// the chart's own part and a copy of `xlsx` as its `externalData`, both written back verbatim on
/// save ([`crate::write::embed`]).
#[allow(clippy::too_many_arguments)]
pub fn insert_chart_from_xlsx(
    doc: &mut Document,
    xlsx: &[u8],
    sheet: Option<&str>,
    range: Option<&str>,
    kind: &str,
    title: Option<&str>,
    w: f32,
    h: f32,
) -> Result<Graphic, DocxError> {
    let grid = xlsx::read_sheet(xlsx, sheet)?;
    let bounds = match range {
        Some(r) => xlsx::parse_range(r).ok_or_else(|| DocxError::Xml(format!("not a cell range: {r:?}")))?,
        None => {
            let (last_row, last_col) = grid.bounds().unwrap_or((0, 0));
            (0, 0, last_col.min(MAX_DEFAULT_COLS - 1), last_row.min(MAX_DEFAULT_ROWS - 1))
        }
    };
    let bounds = clamp_bounds(bounds);
    let data = chart_data(&grid, bounds);
    if data.series.iter().all(|(_, v)| v.iter().all(Option::is_none)) {
        return Err(DocxError::Xml("the selected range has no numeric data to chart".into()));
    }
    let kind = Kind::parse(kind);
    let sheet_name = sheet.unwrap_or("Sheet1");
    let chart_bytes = chart_space_xml(kind, &data, title, sheet_name, bounds);

    let mut manifest = doc.passthrough.get(EMBEDDED_PARTS).map(|b| EmbeddedManifest::parse(b)).unwrap_or_default();
    let chart_part = alloc_part(doc, &manifest, "word/charts/chart", ".xml");
    let xlsx_part = alloc_part(doc, &manifest, "word/embeddings/Microsoft_Excel_Worksheet", ".xlsx");
    manifest.parts.insert(
        chart_part.clone(),
        EmbeddedPart {
            content_type: CHART_CT.to_string(),
            rels: vec![Rel { id: "rId1".to_string(), kind: rt::PACKAGE.to_string(), target: xlsx_part.clone(), external: false }],
        },
    );
    manifest.parts.insert(xlsx_part.clone(), EmbeddedPart { content_type: XLSX_CT.to_string(), rels: Vec::new() });
    doc.passthrough.insert(EMBEDDED_PARTS.to_string(), Arc::new(manifest.to_bytes()));
    doc.passthrough.insert(chart_part.clone(), Arc::new(chart_bytes.clone()));
    doc.passthrough.insert(xlsx_part, Arc::new(xlsx.to_vec()));

    let space = xml::parse(&chart_bytes)?;
    let items = super::chart::chart_items(&space, &doc.settings.theme_colors, w, h);
    let embedded = Embedded {
        xml: graphic_frame_xml(),
        refs: vec![EmbeddedRef { id: "rId1".to_string(), kind: rt::CHART.to_string(), target: chart_part, external: false }],
    };
    Ok(Graphic { kind: GraphicKind::Chart, items, w, h, source: Some(Arc::new(embedded)) })
}

/// A free part name under `{stem}{n}{ext}` (`doc.passthrough` and `manifest` both checked).
fn alloc_part(doc: &Document, manifest: &EmbeddedManifest, stem: &str, ext: &str) -> String {
    for n in 1..=MAX_NAME_ATTEMPTS {
        let name = format!("{stem}{n}{ext}");
        if !doc.passthrough.contains_key(&name) && !manifest.parts.contains_key(&name) {
            return name;
        }
    }
    format!("{stem}{MAX_NAME_ATTEMPTS}{ext}")
}

/// Caps a (col0, row0, col1, row1) selection to the most series and categories a baked chart
/// holds.
fn clamp_bounds((c0, r0, c1, r1): (usize, usize, usize, usize)) -> (usize, usize, usize, usize) {
    (c0, r0, c1.min(c0.saturating_add(MAX_SERIES)), r1.min(r0.saturating_add(MAX_CATEGORIES)))
}

fn chart_data(grid: &Grid, (c0, r0, c1, r1): (usize, usize, usize, usize)) -> ChartData {
    let rows = r1 - r0 + 1;
    let cols = c1 - c0 + 1;
    let text = |r: usize, c: usize| grid.get(r, c).map(|v| v.text.clone()).unwrap_or_default();
    let num = |r: usize, c: usize| grid.get(r, c).and_then(|v| v.number);
    if rows > 1 && cols > 1 {
        let categories: Vec<String> = (r0 + 1..=r1).map(|r| text(r, c0)).collect();
        let series = (c0 + 1..=c1)
            .map(|c| {
                let name = text(r0, c);
                let name = if name.is_empty() { format!("Series {}", c - c0) } else { name };
                (name, (r0 + 1..=r1).map(|r| num(r, c)).collect())
            })
            .collect();
        ChartData { categories, series }
    } else if cols == 1 {
        let categories = (1..=rows).map(|i| i.to_string()).collect();
        ChartData { categories, series: vec![("Series 1".to_string(), (r0..=r1).map(|r| num(r, c0)).collect())] }
    } else {
        let categories = (1..=cols).map(|i| i.to_string()).collect();
        ChartData { categories, series: vec![("Series 1".to_string(), (c0..=c1).map(|c| num(r0, c)).collect())] }
    }
}

/// The `a:graphic` drawing markup pointing at the chart part (relationship `rId1`, rewritten to
/// the document's own id when written — see [`crate::write::embed`]), as a standalone element
/// with the namespace declarations it needs (as [`super::embed::keep_embedded`] would produce
/// for one read from a file).
fn graphic_frame_xml() -> String {
    format!(
        "<a:graphic xmlns:a=\"{a}\" xmlns:c=\"{c}\" xmlns:r=\"{r}\"><a:graphicData uri=\"{c}\"><c:chart r:id=\"rId1\"/></a:graphicData></a:graphic>",
        a = "http://schemas.openxmlformats.org/drawingml/2006/main",
        c = "http://schemas.openxmlformats.org/drawingml/2006/chart",
        r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    )
}

/// A chart part (`c:chartSpace`, ECMA-376 Part 1 §21.2) with `data`'s cache baked in and a
/// `c:externalData` pointing at relationship `rId1` (the embedded workbook).
fn chart_space_xml(kind: Kind, data: &ChartData, title: Option<&str>, sheet: &str, bounds: (usize, usize, usize, usize)) -> Vec<u8> {
    let mut w = W { s: String::new() };
    w.open(
        "c:chartSpace",
        &[
            ("xmlns:c", "http://schemas.openxmlformats.org/drawingml/2006/chart"),
            ("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
            ("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
        ],
    );
    w.empty("c:date1904", &[("val", "0")]);
    w.empty("c:lang", &[("val", "en-US")]);
    w.empty("c:roundedCorners", &[("val", "0")]);
    w.open("c:chart", &[]);
    match title.filter(|t| !t.is_empty()) {
        Some(t) => {
            w.open("c:title", &[]);
            w.open("c:tx", &[]);
            w.open("c:rich", &[]);
            w.empty("a:bodyPr", &[]);
            w.open("a:p", &[]);
            w.open("a:r", &[]);
            w.leaf("a:t", &[], t);
            w.close("a:r");
            w.close("a:p");
            w.close("c:rich");
            w.close("c:tx");
            w.empty("c:overlay", &[("val", "0")]);
            w.close("c:title");
            w.empty("c:autoTitleDeleted", &[("val", "0")]);
        }
        None => w.empty("c:autoTitleDeleted", &[("val", "1")]),
    }
    w.open("c:plotArea", &[]);
    w.empty("c:layout", &[]);
    let f = formula_range(sheet, bounds);
    match kind {
        Kind::Pie | Kind::Doughnut => pie_chart_xml(&mut w, kind, data, &f),
        _ => cartesian_chart_xml(&mut w, kind, data, &f),
    }
    w.close("c:plotArea");
    if data.series.len() > 1 || matches!(kind, Kind::Pie | Kind::Doughnut) {
        w.open("c:legend", &[]);
        w.empty("c:legendPos", &[("val", "b")]);
        w.empty("c:overlay", &[("val", "0")]);
        w.close("c:legend");
    }
    w.empty("c:plotVisOnly", &[("val", "1")]);
    w.empty("c:dispBlanksAs", &[("val", "gap")]);
    w.close("c:chart");
    w.open("c:externalData", &[("r:id", "rId1")]);
    w.empty("c:autoUpdate", &[("val", "0")]);
    w.close("c:externalData");
    w.close("c:chartSpace");
    w.into_bytes()
}

fn cartesian_chart_xml(w: &mut W, kind: Kind, data: &ChartData, f: &str) {
    // `chart_space_xml` only calls this for Bar, Line or Area; Pie and Doughnut go to
    // `pie_chart_xml` instead (never reached here, but falling back to a bar tag is harmless).
    let tag = match kind {
        Kind::Line => "c:lineChart",
        Kind::Area => "c:areaChart",
        Kind::Bar | Kind::Pie | Kind::Doughnut => "c:barChart",
    };
    w.open(tag, &[]);
    if kind == Kind::Bar {
        w.empty("c:barDir", &[("val", "col")]);
    }
    w.empty("c:grouping", &[("val", if kind == Kind::Bar { "clustered" } else { "standard" })]);
    w.empty("c:varyColors", &[("val", "0")]);
    for (idx, (name, values)) in data.series.iter().enumerate() {
        write_series(w, kind, idx, name, values, &data.categories, f);
    }
    if kind == Kind::Bar {
        w.empty("c:gapWidth", &[("val", "150")]);
    }
    w.empty("c:axId", &[("val", "111111111")]);
    w.empty("c:axId", &[("val", "222222222")]);
    w.close(tag);
    w.open("c:catAx", &[]);
    w.empty("c:axId", &[("val", "111111111")]);
    w.open("c:scaling", &[]);
    w.empty("c:orientation", &[("val", "minMax")]);
    w.close("c:scaling");
    w.empty("c:delete", &[("val", "0")]);
    w.empty("c:axPos", &[("val", "b")]);
    w.empty("c:crossAx", &[("val", "222222222")]);
    w.close("c:catAx");
    w.open("c:valAx", &[]);
    w.empty("c:axId", &[("val", "222222222")]);
    w.open("c:scaling", &[]);
    w.empty("c:orientation", &[("val", "minMax")]);
    w.close("c:scaling");
    w.empty("c:delete", &[("val", "0")]);
    w.empty("c:axPos", &[("val", "l")]);
    w.empty("c:majorGridlines", &[]);
    w.empty("c:crossAx", &[("val", "111111111")]);
    w.close("c:valAx");
}

fn pie_chart_xml(w: &mut W, kind: Kind, data: &ChartData, f: &str) {
    let tag = if kind == Kind::Doughnut { "c:doughnutChart" } else { "c:pieChart" };
    w.open(tag, &[]);
    w.empty("c:varyColors", &[("val", "1")]);
    if let Some((name, values)) = data.series.first() {
        write_series(w, kind, 0, name, values, &data.categories, f);
    }
    w.empty("c:firstSliceAng", &[("val", "0")]);
    if kind == Kind::Doughnut {
        w.empty("c:holeSize", &[("val", "50")]);
    }
    w.close(tag);
}

fn write_series(w: &mut W, kind: Kind, idx: usize, name: &str, values: &[Option<f64>], categories: &[String], f: &str) {
    w.open("c:ser", &[]);
    w.empty("c:idx", &[("val", &idx.to_string())]);
    w.empty("c:order", &[("val", &idx.to_string())]);
    w.open("c:tx", &[]);
    w.leaf("c:v", &[], name);
    w.close("c:tx");
    if kind == Kind::Line {
        w.open("c:marker", &[]);
        w.empty("c:symbol", &[("val", "none")]);
        w.close("c:marker");
    }
    write_str_ref(w, "c:cat", f, categories);
    write_num_ref(w, "c:val", f, values);
    if kind == Kind::Line {
        w.empty("c:smooth", &[("val", "0")]);
    }
    w.close("c:ser");
}

fn write_str_ref(w: &mut W, tag: &str, f: &str, items: &[String]) {
    w.open(tag, &[]);
    w.open("c:strRef", &[]);
    w.leaf("c:f", &[], f);
    w.open("c:strCache", &[]);
    w.empty("c:ptCount", &[("val", &items.len().to_string())]);
    for (i, s) in items.iter().enumerate().filter(|(_, s)| !s.is_empty()) {
        w.open("c:pt", &[("idx", &i.to_string())]);
        w.leaf("c:v", &[], s);
        w.close("c:pt");
    }
    w.close("c:strCache");
    w.close("c:strRef");
    w.close(tag);
}

fn write_num_ref(w: &mut W, tag: &str, f: &str, values: &[Option<f64>]) {
    w.open(tag, &[]);
    w.open("c:numRef", &[]);
    w.leaf("c:f", &[], f);
    w.open("c:numCache", &[]);
    w.leaf("c:formatCode", &[], "General");
    w.empty("c:ptCount", &[("val", &values.len().to_string())]);
    for (i, v) in values.iter().enumerate().filter_map(|(i, v)| (*v).map(|v| (i, v))) {
        w.open("c:pt", &[("idx", &i.to_string())]);
        w.leaf("c:v", &[], &format!("{v}"));
        w.close("c:pt");
    }
    w.close("c:numCache");
    w.close("c:numRef");
    w.close(tag);
}

/// A formula string for `c:f`: the whole selection, quoted with its sheet name. Every series and
/// category reference points at it (the chart isn't a live link to the workbook — baking the
/// cache once is this command's whole job — so a precise per-column reference buys nothing).
fn formula_range(sheet: &str, (c0, r0, c1, r1): (usize, usize, usize, usize)) -> String {
    format!("'{}'!${}${}:${}${}", sheet.replace('\'', "''"), xlsx::col_letters(c0), r0 + 1, xlsx::col_letters(c1), r1 + 1)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn grid(rows: &[&[&str]]) -> Grid {
        let xml = sheet_xml(rows);
        let el = xml::parse(xml.as_bytes()).unwrap();
        crate::xlsx::parse_sheet(&el, &[])
    }

    /// A minimal `sheetData` XML for a grid of string cells (inline strings, so no shared-string
    /// table is needed).
    fn sheet_xml(rows: &[&[&str]]) -> String {
        let mut s = String::from(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#);
        for (r, row) in rows.iter().enumerate() {
            s.push_str(&format!(r#"<row r="{}">"#, r + 1));
            for (c, v) in row.iter().enumerate() {
                let col = xlsx::col_letters(c);
                let r1 = r + 1;
                if let Ok(n) = v.parse::<f64>() {
                    s.push_str(&format!(r#"<c r="{col}{r1}"><v>{n}</v></c>"#));
                } else if !v.is_empty() {
                    s.push_str(&format!(r#"<c r="{col}{r1}" t="inlineStr"><is><t>{v}</t></is></c>"#));
                }
            }
            s.push_str("</row>");
        }
        s.push_str("</sheetData></worksheet>");
        s
    }

    #[test]
    fn header_row_and_column_become_series_and_categories() {
        let g = grid(&[&["", "Q1", "Q2"], &["Jan", "1", "2"], &["Feb", "3", "4"]]);
        let data = chart_data(&g, (0, 0, 2, 2));
        assert_eq!(data.categories, vec!["Jan", "Feb"]);
        assert_eq!(data.series, vec![("Q1".to_string(), vec![Some(1.0), Some(3.0)]), ("Q2".to_string(), vec![Some(2.0), Some(4.0)])]);
    }

    #[test]
    fn a_single_column_is_one_series_by_position() {
        let g = grid(&[&["5"], &["6"], &["7"]]);
        let data = chart_data(&g, (0, 0, 0, 2));
        assert_eq!(data.categories, vec!["1", "2", "3"]);
        assert_eq!(data.series, vec![("Series 1".to_string(), vec![Some(5.0), Some(6.0), Some(7.0)])]);
    }

    #[test]
    fn bar_chart_xml_is_well_formed_and_carries_the_cache() {
        let data = ChartData { categories: vec!["Jan".into(), "Feb".into()], series: vec![("Q1".into(), vec![Some(1.0), Some(3.0)])] };
        let bytes = chart_space_xml(Kind::Bar, &data, Some("Sales"), "Sheet1", (0, 0, 1, 2));
        let space = xml::parse(&bytes).unwrap();
        assert!(space.child("c:chart").unwrap().child("c:title").is_some());
        let items = crate::read::chart::chart_items(&space, &wordcraft_doc::THEME_COLORS, 300.0, 200.0);
        assert!(!items.is_empty(), "a bar chart draws something");
        assert!(bytes.windows(b"$A$1:$B$3".len()).any(|w| w == b"$A$1:$B$3"), "{}", String::from_utf8_lossy(&bytes));
    }

    #[test]
    fn pie_chart_xml_keeps_only_the_first_series() {
        let data = ChartData {
            categories: vec!["A".into(), "B".into()],
            series: vec![("S1".into(), vec![Some(1.0), Some(2.0)]), ("S2".into(), vec![Some(3.0), Some(4.0)])],
        };
        let bytes = chart_space_xml(Kind::Pie, &data, None, "Sheet1", (0, 0, 2, 2));
        let space = xml::parse(&bytes).unwrap();
        let pie = space.child("c:chart").unwrap().child("c:plotArea").unwrap().child("c:pieChart").unwrap();
        assert_eq!(pie.children("c:ser").count(), 1);
    }

    #[test]
    fn clamp_bounds_caps_a_hostile_range() {
        let (c0, r0, c1, r1) = clamp_bounds((0, 0, usize::MAX, usize::MAX));
        assert_eq!((c0, r0), (0, 0));
        assert!(c1 - c0 <= MAX_SERIES && r1 - r0 <= MAX_CATEGORIES);
    }

    /// A minimal real `.xlsx` package: one sheet, a shared string for the header, numeric data.
    fn xlsx_bytes() -> Vec<u8> {
        let workbook = br#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
        let workbook_rels = br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#;
        let shared_strings =
            br#"<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Units</t></si></sst>"#;
        let sheet1 = br#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="B1" t="s"><v>0</v></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Jan</t></is></c><c r="B2"><v>10</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Feb</t></is></c><c r="B3"><v>20</v></c></row></sheetData></worksheet>"#;
        let mut zw = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        for (name, bytes) in [
            ("xl/workbook.xml", workbook.as_slice()),
            ("xl/_rels/workbook.xml.rels", workbook_rels.as_slice()),
            ("xl/sharedStrings.xml", shared_strings.as_slice()),
            ("xl/worksheets/sheet1.xml", sheet1.as_slice()),
        ] {
            zw.start_file(name, zip::write::SimpleFileOptions::default()).unwrap();
            std::io::Write::write_all(&mut zw, bytes).unwrap();
        }
        zw.finish().unwrap().into_inner()
    }

    #[test]
    fn inserting_a_real_xlsx_bakes_a_chart_and_keeps_the_workbook() {
        let mut doc = wordcraft_doc::Document::new();
        let xlsx = xlsx_bytes();
        let graphic = insert_chart_from_xlsx(&mut doc, &xlsx, None, None, "bar", Some("Units sold"), 300.0, 200.0).unwrap();
        assert!(!graphic.items.is_empty());
        let src = graphic.source.as_ref().unwrap();
        assert_eq!(src.refs.len(), 1);
        let chart_part = &src.refs[0].target;
        let chart_bytes = doc.passthrough.get(chart_part).unwrap();
        let space = xml::parse(chart_bytes).unwrap();
        let ext = space.child("c:externalData").unwrap();
        let xlsx_rid = ext.attr("r:id").unwrap();

        let manifest = crate::package::EmbeddedManifest::parse(doc.passthrough.get(crate::package::EMBEDDED_PARTS).unwrap());
        let chart_entry = manifest.parts.get(chart_part).unwrap();
        let xlsx_rel = chart_entry.rels.iter().find(|r| r.id == xlsx_rid).unwrap();
        assert!(doc.passthrough.contains_key(&xlsx_rel.target));
        assert_eq!(doc.passthrough.get(&xlsx_rel.target).unwrap().as_slice(), xlsx.as_slice());
    }

    #[test]
    fn a_bad_xlsx_is_an_error_not_a_panic() {
        let mut doc = wordcraft_doc::Document::new();
        assert!(insert_chart_from_xlsx(&mut doc, b"not a zip", None, None, "bar", None, 300.0, 200.0).is_err());
        assert!(insert_chart_from_xlsx(&mut doc, b"", None, None, "bar", None, 300.0, 200.0).is_err());
    }
}
