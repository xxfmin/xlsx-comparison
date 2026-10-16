package parity;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.UnaryOperator;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * One case per schema gap found by walking the ECMA-376 schema (G1-G18). Each case edits a real
 * workbook's XML and asserts the comparator puts the change in the right tier: DIFFERENCE fails
 * the comparison, INFO is reported without failing it, NONE (an encoding-only change) must
 * compare strictly identical.
 */
class GapCoverageTest {

    enum Tier { DIFFERENCE, INFO, NONE }

    @TempDir static Path dir;
    static Path base;
    static final AtomicInteger seq = new AtomicInteger();

    @BeforeAll
    static void buildBase() throws Exception {
        base = Fixtures.statement(dir.resolve("base.xlsx"));
    }

    // ------------------------------------------------------------ variant builder

    /** A set of edits and additions applied to the base workbook. */
    static final class V {
        final Map<String, UnaryOperator<String>> edits = new LinkedHashMap<>();
        final Map<String, byte[]> adds = new LinkedHashMap<>();

        V edit(String entry, UnaryOperator<String> op) {
            edits.merge(entry, op, (a, b) -> s -> b.apply(a.apply(s)));
            return this;
        }

        V replace(String entry, String from, String to) {
            return edit(entry, s -> s.replace(from, to));
        }

        V regex(String entry, String pattern, String to) {
            return edit(entry, s -> s.replaceFirst(pattern, to));
        }

        V add(String entry, String content) {
            adds.put(entry, content.getBytes(StandardCharsets.UTF_8));
            return this;
        }

        V add(String entry, byte[] content) {
            adds.put(entry, content);
            return this;
        }

        V contentType(String xml) { return replace("[Content_Types].xml", "</Types>", xml + "</Types>"); }

        V sheet1Rel(String id, String type, String target) {
            return replace("xl/worksheets/_rels/sheet1.xml.rels", "</Relationships>",
                "<Relationship Id=\"" + id + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                    + type + "\" Target=\"" + target + "\"/></Relationships>");
        }

        /** Adds an Office theme, optionally edited; POI-created workbooks have none. */
        V themed(UnaryOperator<String> themeEdit) {
            add("xl/theme/theme1.xml", themeEdit.apply(THEME));
            contentType("<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>");
            return replace("xl/_rels/workbook.xml.rels", "</Relationships>",
                "<Relationship Id=\"rIdTheme\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/></Relationships>");
        }

        V background(int rgb) throws Exception {
            replace("xl/worksheets/sheet1.xml", "<tableParts", "<picture r:id=\"rIdBg\"/><tableParts");
            sheet1Rel("rIdBg", "image", "../media/bg.png");
            return add("xl/media/bg.png", Fixtures.png(rgb));
        }

        V headerImage(int rgb) throws Exception {
            replace("xl/worksheets/sheet1.xml", "<tableParts", "<legacyDrawingHF r:id=\"rIdHF\"/><tableParts");
            sheet1Rel("rIdHF", "vmlDrawing", "../drawings/vmlDrawingHF.vml");
            add("xl/drawings/vmlDrawingHF.vml", "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\">"
                + "<v:shape id=\"CH\" style=\"position:absolute;width:20pt;height:20pt\"><v:imagedata o:relid=\"rId1\" o:title=\"logo\"/></v:shape></xml>");
            add("xl/drawings/_rels/vmlDrawingHF.vml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/hf.png\"/></Relationships>");
            return add("xl/media/hf.png", Fixtures.png(rgb));
        }

        /** Two header images (left and right) whose shapes point at images by o:relid. */
        V headerLogos(String leftRel, String rightRel) throws Exception {
            replace("xl/worksheets/sheet1.xml", "<tableParts", "<legacyDrawingHF r:id=\"rIdHF\"/><tableParts");
            sheet1Rel("rIdHF", "vmlDrawing", "../drawings/vmlDrawingHF.vml");
            add("xl/drawings/vmlDrawingHF.vml", "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\">"
                + "<v:shape id=\"LH\" style=\"position:absolute;width:20pt;height:20pt\"><v:imagedata o:relid=\"" + leftRel + "\" o:title=\"left\"/></v:shape>"
                + "<v:shape id=\"RH\" style=\"position:absolute;width:20pt;height:20pt\"><v:imagedata o:relid=\"" + rightRel + "\" o:title=\"right\"/></v:shape></xml>");
            add("xl/drawings/_rels/vmlDrawingHF.vml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/logoA.png\"/>"
                + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/logoB.png\"/></Relationships>");
            add("xl/media/logoA.png", Fixtures.png(0xFF0000));
            return add("xl/media/logoB.png", Fixtures.png(0x0000FF));
        }

        V ole(String content) {
            contentType("<Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.oleObject\"/>");
            replace("xl/worksheets/sheet1.xml", "<tableParts",
                "<oleObjects><oleObject progId=\"Package\" shapeId=\"1025\" r:id=\"rIdOle\"/></oleObjects><tableParts");
            sheet1Rel("rIdOle", "oleObject", "../embeddings/oleObject1.bin");
            return add("xl/embeddings/oleObject1.bin", content);
        }

        V drawing1(String anchorXml) {
            return replace("xl/drawings/drawing1.xml", "</xdr:wsDr>", anchorXml + "</xdr:wsDr>");
        }

        Path build() throws Exception {
            if (edits.isEmpty() && adds.isEmpty()) return base;
            return Fixtures.rewrite(base, dir.resolve("v" + seq.incrementAndGet() + ".xlsx"), edits, adds);
        }
    }

    static V v() { return new V(); }

    record Case(String name, V expected, V actual, Tier tier, String fragment) {
        @Override public String toString() { return name; }
    }

    static Case diff(String name, V actual, String fragment) { return new Case(name, v(), actual, Tier.DIFFERENCE, fragment); }
    static Case info(String name, V actual, String fragment) { return new Case(name, v(), actual, Tier.INFO, fragment); }
    static Case none(String name, V actual) { return new Case(name, v(), actual, Tier.NONE, ""); }

    // ------------------------------------------------------------ the gaps

    static final String S1 = "xl/worksheets/sheet1.xml", S2 = "xl/worksheets/sheet2.xml", STYLES = "xl/styles.xml",
        DRAW = "xl/drawings/drawing1.xml", WB = "xl/workbook.xml";

    static String anchor(int col, int row, String shapeXml) {
        return "<xdr:twoCellAnchor><xdr:from><xdr:col>" + col + "</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>" + row
            + "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>" + (col + 2) + "</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>"
            + (row + 3) + "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" + shapeXml + "<xdr:clientData/></xdr:twoCellAnchor>";
    }

    static String group(int childX) {
        return anchor(10, 0, "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"10\" name=\"Group 1\"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>"
            + "<xdr:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100\" cy=\"100\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"100\" cy=\"100\"/></a:xfrm></xdr:grpSpPr>"
            + "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"11\" name=\"Child\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"" + childX
            + "\" y=\"0\"/><a:ext cx=\"50\" cy=\"50\"/></a:xfrm><a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp></xdr:grpSp>");
    }

    static final String DXF1 = "<dxfs count=\"1\"><dxf><font><color rgb=\"FF0000\"/></font></dxf></dxfs>";

    static String customTableStyle(String headerFill) {
        return "<dxfs count=\"2\"><dxf><font><color rgb=\"FF0000\"/></font></dxf><dxf><fill><patternFill><bgColor rgb=\"FF" + headerFill
            + "\"/></patternFill></fill></dxf></dxfs><tableStyles count=\"1\"><tableStyle name=\"Bank\" count=\"1\">"
            + "<tableStyleElement type=\"headerRow\" dxfId=\"1\"/></tableStyle></tableStyles>";
    }

    static Stream<Case> cases() throws Exception {
        List<Case> c = new ArrayList<>();
        // G1 phonetic guides
        c.add(diff("G1 cell shows phonetic guide", v().replace(S1, "<c r=\"A2\" ", "<c r=\"A2\" ph=\"1\" "), "S1!A2"));
        c.add(diff("G1 row shows phonetic guide", v().replace(S1, "<row r=\"3\">", "<row r=\"3\" ph=\"1\">"), "S1.row[3]"));
        c.add(diff("G1 phonetic guide text", v().replace("xl/sharedStrings.xml", "<si><t>Opening</t></si>",
            "<si><t>Opening</t><rPh sb=\"0\" eb=\"2\"><t>オ</t></rPh></si>"), "S1!A2"));
        // G2 formula flags
        c.add(diff("G2 formula assigns to name (bx)", v().replace(S1, "<f>B2*2</f>", "<f bx=\"1\">B2*2</f>"), "S1!C2"));
        c.add(diff("G2 what-if data table formula", v().replace(S1, "<f>B2*2</f>", "<f t=\"dataTable\" ref=\"C2\" dt2D=\"0\" dtr=\"1\" r1=\"B2\"/>"), "S1!C2"));
        c.add(info("G2 calculate-always flag", v().replace(S1, "<f>B2*2</f>", "<f ca=\"1\">B2*2</f>"), "info.S1!C2"));
        // G3 cell extension list
        c.add(diff("G3 cell extension list", v().replace(S1, "<v>1250.5</v></c>",
            "<v>1250.5</v><extLst><ext uri=\"{0E5E9B5D-1A2B-4C3D-8E9F-000000000001}\"/></extLst></c>"), "S1!B2"));
        // G4 row height re-fit behaviour
        c.add(info("G4 row height custom flag", v().replace(S1, " customHeight=\"true\"", ""), "info.S1.rowCustomHeight[4]"));
        // G5 conditional formatting
        c.add(diff("G5 highlight colour (resolved dxf)", v().replace(STYLES, "<color rgb=\"FF0000\"/></font></dxf>",
            "<color rgb=\"00B050\"/></font></dxf>"), "S1.cf[B2:B5]"));
        c.add(diff("G5 rule option stopIfTrue", v().replace(S1, "<cfRule type=\"cellIs\"", "<cfRule stopIfTrue=\"1\" type=\"cellIs\""), "S1.cf[B2:B5]"));
        c.add(diff("G5 colour scale colour", v().replace(S2, "F8696B", "000000"), "S2.cf[C2:C6]"));
        c.add(none("G5 priority renumbered, same order", v().replace(S1, "priority=\"1\"", "priority=\"7\"")));
        // G6 data validation
        c.add(diff("G6 input method mode", v().replace(S1, "<dataValidation type=\"list\"", "<dataValidation imeMode=\"hiragana\" type=\"list\""), "S1.validation[D2:D5]"));
        c.add(diff("G6 prompts disabled for sheet", v().replace(S1, "<dataValidations count=\"1\">", "<dataValidations count=\"1\" disablePrompts=\"1\">"), "S1.validations"));
        c.add(none("G6 count attribute only", v().replace(S1, "<dataValidations count=\"1\">", "<dataValidations count=\"9\">")));
        // G7 print options and page breaks
        c.add(diff("G7 gridLinesSet", v().replace(S1, "<printOptions horizontalCentered=\"true\"/>",
            "<printOptions horizontalCentered=\"true\" gridLinesSet=\"0\"/>"), "S1.printOptionsXml"));
        c.add(diff("G7 page break extent", v().replace(S1, "max=\"16383\"", "max=\"5\""), "S1.rowBreaks"));
        c.add(none("G7 break count attributes only", v().replace(S1, "manualBreakCount=\"1\"", "manualBreakCount=\"4\"")));
        // G8 drawings
        c.add(diff("G8 shape rotation", v().regex(DRAW, "(<xdr:sp>.*?<a:xfrm)>", "$1 rot=\"5400000\">"), "S1.shape@"));
        c.add(diff("G8 picture alt text", v().replace(DRAW, "descr=\"Picture\"", "descr=\"Bank logo\""), "S1.picture@"));
        c.add(diff("G8 picture crop", v().replace(DRAW, "<a:blip r:embed=\"rId1\"/>", "<a:blip r:embed=\"rId1\"/><a:srcRect l=\"10000\"/>"), "S1.picture@"));
        c.add(diff("G8 connector added", v().drawing1(anchor(10, 6, "<xdr:cxnSp macro=\"\"><xdr:nvCxnSpPr><xdr:cNvPr id=\"12\" name=\"Connector 1\"/>"
            + "<xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:prstGeom prst=\"line\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:cxnSp>")), "S1.connector@"));
        c.add(diff("G8 absolute-anchored shape", v().drawing1("<xdr:absoluteAnchor><xdr:pos x=\"100\" y=\"200\"/><xdr:ext cx=\"300\" cy=\"400\"/>"
            + "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"13\" name=\"Abs\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"rect\">"
            + "<a:avLst/></a:prstGeom></xdr:spPr></xdr:sp><xdr:clientData/></xdr:absoluteAnchor>"), "S1.shape@abs:"));
        c.add(new Case("G8 shape moved inside a group", v().drawing1(group(0)), v().drawing1(group(40)), Tier.DIFFERENCE, "S1.group@"));
        c.add(diff("G8 chart title", v().replace("xl/charts/chart1.xml", "<a:t>Balance</a:t>", "<a:t>Balances</a:t>"), "S2.frame@"));
        c.add(diff("G8 chart bar direction", v().replace("xl/charts/chart1.xml", "<c:barDir val=\"col\"/>", "<c:barDir val=\"bar\"/>"), "S2.frame@"));
        c.add(info("G8 shape renamed", v().replace(DRAW, "name=\"Shape 1\"", "name=\"Callout\""), "info.S1.shape@"));
        c.add(none("G8 shape id renumbered", v().replace(DRAW, "<xdr:cNvPr id=\"2\" ", "<xdr:cNvPr id=\"42\" ")));
        // G9 comments
        c.add(diff("G9 comment box geometry", v().replace("xl/drawings/vmlDrawing0.vml", "2, 0, 1, 0, 4, 0, 4, 0", "2, 0, 1, 0, 6, 0, 8, 0"), "S1.comment[B2]"));
        c.add(diff("G9 comment fill colour", v().replace("xl/drawings/vmlDrawing0.vml", "fillcolor=\"#ffffe1\"", "fillcolor=\"#ff0000\""), "S1.comment[B2]"));
        c.add(diff("G9 comment text formatting", v().replace("xl/comments1.xml", "<text><t>Includes pending items</t></text>",
            "<text><r><rPr><b/></rPr><t>Includes pending items</t></r></text>"), "S1.comment[B2]"));
        // G10 tables
        c.add(diff("G10 table banding", v().replace("xl/tables/table1.xml", "showRowStripes=\"true\"", "showRowStripes=\"false\""), "S1.table[Txns]"));
        c.add(diff("G10 totals row function", v().replace("xl/tables/table1.xml", "<tableColumn id=\"3\" name=\"Col2\"/>",
            "<tableColumn id=\"3\" name=\"Col2\" totalsRowFunction=\"sum\"/>"), "S1.table[Txns]"));
        c.add(none("G10 table id renumbered", v().replace("xl/tables/table1.xml", "<table id=\"1\"", "<table id=\"5\"")));
        // G11 content in other parts
        c.add(diff("G11 background image added", v().background(0xFF0000), "S1.xml<picture>"));
        c.add(new Case("G11 background image content", v().background(0xFF0000), v().background(0x00FF00), Tier.DIFFERENCE, "S1.xml<picture>->1"));
        c.add(new Case("G11 header image content (VML o:relid)", v().headerImage(0xFF0000), v().headerImage(0x0000FF), Tier.DIFFERENCE, "S1.xml<legacyDrawingHF>->1"));
        c.add(new Case("G11 header logos swapped left/right", v().headerLogos("rId1", "rId2"), v().headerLogos("rId2", "rId1"),
            Tier.DIFFERENCE, "S1.xml<legacyDrawingHF>->1"));
        c.add(new Case("G11 embedded OLE object content", v().ole("first"), v().ole("second"), Tier.DIFFERENCE, "S1.xml<oleObjects>->1"));
        // G12 styles beyond cell formats
        c.add(new Case("G12 named style renamed",
            v().replace(STYLES, "</cellXfs>", "</cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"),
            v().replace(STYLES, "</cellXfs>", "</cellXfs><cellStyles count=\"1\"><cellStyle name=\"Standard\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"),
            Tier.INFO, "info."));
        c.add(new Case("G12 custom table style colour",
            v().replace(STYLES, DXF1, customTableStyle("4472C4")).replace("xl/tables/table1.xml", "TableStyleMedium2", "Bank"),
            v().replace(STYLES, DXF1, customTableStyle("C00000")).replace("xl/tables/table1.xml", "TableStyleMedium2", "Bank"),
            Tier.DIFFERENCE, "S1.table[Txns]"));
        c.add(info("G12 recently used colours", v().replace(STYLES, "</dxfs>", "</dxfs><colors><mruColors><color rgb=\"FF123456\"/></mruColors></colors>"), "info.styles.other"));
        c.add(diff("G12 custom indexed palette", v().replace(STYLES, "</dxfs>", "</dxfs><colors><indexedColors><rgbColor rgb=\"FF000000\"/></indexedColors></colors>"), "wb.indexedColors"));
        // G13 theme
        c.add(new Case("G13 theme line style", v().themed(t -> t), v().themed(t -> t.replace("w=\"12700\"", "w=\"25400\"")), Tier.DIFFERENCE, "wb.themeFormat.lnStyleLst"));
        c.add(new Case("G13 theme font", v().themed(t -> t), v().themed(t -> t.replace("typeface=\"Calibri\"/>", "typeface=\"Arial\"/>")), Tier.DIFFERENCE, "wb.themeFonts"));
        c.add(new Case("G13 theme colour", v().themed(t -> t), v().themed(t -> t.replace("4472C4", "FF0000")), Tier.DIFFERENCE, "wb.themeColors"));
        c.add(new Case("G13 theme name", v().themed(t -> t), v().themed(t -> t.replace("name=\"Office Theme\"", "name=\"Bank Theme\"")), Tier.INFO, "info.wb.theme"));
        // G14 other parts
        c.add(diff("G14 custom XML part added", v().add("customXml/item1.xml", "<data xmlns=\"urn:test\"><v>1</v></data>"), "part[/customXml/item1.xml]"));
        c.add(new Case("G14 custom XML content (localized)", v().add("customXml/item1.xml", "<data xmlns=\"urn:test\"><v>1</v></data>"),
            v().add("customXml/item1.xml", "<data xmlns=\"urn:test\"><v>2</v></data>"), Tier.DIFFERENCE, "part[/customXml/item1.xml]/v"));
        c.add(new Case("G14 unparseable part compared by hash", v().add("customXml/item1.xml", "<broken"), v().add("customXml/item1.xml", "<broken2"),
            Tier.DIFFERENCE, "part[/customXml/item1.xml]"));
        c.add(new Case("G14 VBA project", vba("macro A"), vba("macro B"), Tier.DIFFERENCE, "part[/xl/vbaProject.bin]"));
        // G15 document properties and derived parts
        c.add(info("G15 document title", v().replace("docProps/core.xml", "<dc:title>Statement</dc:title>", "<dc:title>Statement 2</dc:title>"), "info.docProps.core"));
        c.add(diff("G15 classification label", v().replace("docProps/custom.xml", ">Internal<", ">Public<"), "docProps.custom[Classification]"));
        c.add(info("G15 calculation chain", v().contentType("<Override PartName=\"/xl/calcChain.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml\"/>")
            .add("xl/calcChain.xml", "<calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><c r=\"C2\" i=\"1\"/></calcChain>"), "info.part[/xl/calcChain.xml]"));
        // G16 bookkeeping
        c.add(info("G16 window width", v().replace(WB, "<workbookView activeTab=\"0\"/>", "<workbookView activeTab=\"0\" windowWidth=\"20000\"/>"), "info.wb.window0"));
        c.add(info("G16 selection range", v().replace(S2, "<selection pane=\"bottomLeft\"/>", "<selection pane=\"bottomLeft\" sqref=\"B3:B4\"/>"), "info.S2.sheetView0"));
        c.add(info("G16 file version stamp", v().replace(WB, "<workbookPr", "<fileVersion appName=\"xl\" lastEdited=\"7\"/><workbookPr"), "info.wb.xml<fileVersion>"));
        c.add(info("G16 used-range cache", v().replace(S1, "<dimension ref=\"A1:D13\"/>", "<dimension ref=\"A1:Z99\"/>"), "info.S1.dimension"));
        c.add(info("G16 column bestFit hint", v().replace(S1, "customWidth=\"true\"/>", "customWidth=\"true\" bestFit=\"1\"/>"), "info.S1.colXml[1:1]"));
        c.add(info("G16 printer DPI", v().regex(S1, "(<pageMargins[^>]*/>)", "$1<pageSetup horizontalDpi=\"300\" verticalDpi=\"300\"/>"), "info.S1.pageSetupXml"));
        c.add(info("G16 VBA code name", v().replace(S1, "<dimension", "<sheetPr codeName=\"Sheet1\"/><dimension"), "info.S1.sheetPr"));
        c.add(info("G16 workbook extension block", v().replace(WB, "</workbook>", "<extLst><ext uri=\"{140A7094-0E35-4892-8432-C4D2E57EDEB5}\"/></extLst></workbook>"), "info.wb.xml<extLst>"));
        // G17 font substitution hints
        c.add(info("G17 font family", v().replace(STYLES, "<family val=\"2\"/>", "<family val=\"3\"/>"), "info.style["));
        return c.stream();
    }

    static V vba(String content) {
        return v().contentType("<Default Extension=\"bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>").add("xl/vbaProject.bin", content);
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("cases")
    void everyGapLandsInTheRightTier(Case c) throws Exception {
        Path expected = c.expected().build(), actual = c.actual().build();
        WorkbookComparator.Result r = WorkbookComparator.compare(expected, actual);
        String report = r.report(10);
        switch (c.tier()) {
            case DIFFERENCE -> {
                assertFalse(r.identical(), "real difference reported identical:\n" + report);
                assertTrue(r.differences().stream().anyMatch(d -> d.contains(c.fragment())), "not localized to " + c.fragment() + ":\n" + report);
            }
            case INFO -> {
                assertTrue(r.identical(), "bookkeeping must not fail the default comparison:\n" + report);
                assertFalse(r.identicalStrict(), "bookkeeping must still be reported:\n" + report);
                assertTrue(r.info().stream().anyMatch(d -> d.contains(c.fragment())), "not localized to " + c.fragment() + ":\n" + report);
            }
            case NONE -> assertTrue(r.identicalStrict(), "encoding-only change reported:\n" + report);
        }
    }

    // ------------------------------------------------------------ G18 salted protection hashes

    static Path protectedCopy(String name, String password) throws Exception {
        try (var in = java.nio.file.Files.newInputStream(base); var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(in)) {
            wb.getSheetAt(0).setSheetPassword(password, org.apache.poi.poifs.crypt.HashAlgorithm.sha512); // random salt
            Path f = dir.resolve(name);
            try (var out = java.nio.file.Files.newOutputStream(f)) { wb.write(out); }
            return f;
        }
    }

    @Test
    void saltedHashesAreInfoAndPasswordsAreVerifiedWhenSupplied() throws Exception {
        Path a = protectedCopy("pa.xlsx", "secret"), b = protectedCopy("pb.xlsx", "secret"), c = protectedCopy("pc.xlsx", "other");
        var plain = WorkbookComparator.compare(a, b);
        assertTrue(plain.identical(), plain.report(10));
        assertTrue(plain.info().stream().anyMatch(d -> d.contains("info.S1.protectionXml")), "salt change must be reported: " + plain.report(10));

        var options = new WorkbookSnapshot.Options(List.of("secret", "other"));
        assertTrue(WorkbookComparator.compare(a, b, options).identical(), "same password, different salt");
        var changed = WorkbookComparator.compare(a, c, options);
        assertTrue(changed.differences().stream().anyMatch(d -> d.contains("S1.password")), changed.report(10));
    }

    // ------------------------------------------------------------ theme

    static final String THEME = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements>\
        <a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>\
        <a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1>\
        <a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>\
        <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink>\
        <a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>\
        <a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>\
        <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>\
        <a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>\
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>\
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>\
        </a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/>\
        </a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>\
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>""";
}
