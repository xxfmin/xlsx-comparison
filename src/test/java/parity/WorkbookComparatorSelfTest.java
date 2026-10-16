package parity;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;
import java.util.function.Consumer;
import java.util.function.UnaryOperator;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.*;

class WorkbookComparatorSelfTest {

    @TempDir Path dir;

    @Test void identicalRegenerationPasses() throws Exception {
        assertIdentical(build("a", false, wb -> {}), build("b", false, wb -> {}));
    }

    @Test void styleTableOrderIsIrrelevant() throws Exception {
        assertIdentical(build("a", false, wb -> {}), build("b", true, wb -> {}));
    }

    @Test void explicitBlankThatLooksAbsentIsIgnored() throws Exception {
        assertIdentical(build("a", false, wb -> {}),
            build("b", false, wb -> wb.getSheetAt(1).createRow(40).createCell(6)));
    }

    @Test void detectsValueChange() throws Exception {
        assertDetects("S2!C3", wb -> wb.getSheetAt(1).getRow(2).getCell(2).setCellValue(42.0));
    }

    @Test void detectsFillColorChange() throws Exception {
        assertDetects("S2!A1", wb -> {
            XSSFCellStyle s = wb.createCellStyle();
            s.cloneStyleFrom(wb.getSheetAt(1).getRow(0).getCell(0).getCellStyle());
            s.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0xFF, 0, 0}, null));
            wb.getSheetAt(1).getRow(0).getCell(0).setCellStyle(s);
        });
    }

    @Test void detectsColumnWidthChange() throws Exception {
        assertDetects("S2.cols", wb -> wb.getSheetAt(1).setColumnWidth(1, 9000));
    }

    @Test void detectsRowHeightChange() throws Exception {
        assertDetects("S2.row[2]", wb -> wb.getSheetAt(1).getRow(1).setHeightInPoints(30));
    }

    @Test void detectsMergeChange() throws Exception {
        assertDetects("S2.merge", wb -> wb.getSheetAt(1).addMergedRegion(new CellRangeAddress(10, 10, 0, 2)));
    }

    @Test void detectsPrintSetupChange() throws Exception {
        assertDetects("S1.print", wb -> wb.getSheetAt(0).getPrintSetup().setLandscape(true));
    }

    @Test void detectsSheetReorder() throws Exception {
        assertDetects("S1.sheet", wb -> wb.setSheetOrder("Transactions", 0));
    }

    @Test void detectsPictureChange() throws Exception {
        assertDetects("S1.picture", wb -> {
            XSSFPicture pic = (XSSFPicture) wb.getSheetAt(0).getDrawingPatriarch().getShapes().stream()
                .filter(XSSFPicture.class::isInstance).findFirst().orElseThrow();
            try (OutputStream o = pic.getPictureData().getPackagePart().getOutputStream()) {
                o.write(png(0x00FF00));
            } catch (IOException e) { throw new UncheckedIOException(e); }
        });
    }

    @Test void colourScaleIsComparedNotJustFlagged() throws Exception {
        Path a = build("a", false, wb -> {}), b = build("b", false, wb -> {
            XSSFSheetConditionalFormatting cf = wb.getSheetAt(1).getSheetConditionalFormatting();
            cf.addConditionalFormatting(new CellRangeAddress[]{CellRangeAddress.valueOf("C2:C4")},
                cf.createConditionalFormattingColorScaleRule());
        });
        var r = WorkbookComparator.compare(a, b);
        assertTrue(r.differences().stream().anyMatch(d -> d.contains("S2.cf[C2:C4]")), r.report(20));
        assertTrue(r.differences().stream().noneMatch(d -> d.startsWith("UNREADABLE")), r.report(20));
    }

    @Test void corruptPartFailsClosed() throws Exception {
        Path a = build("a", false, wb -> {});
        Path broken = dir.resolve("broken.xlsx");
        rewriteEntry(a, broken, "xl/drawings/drawing1.xml", x -> x.replace("</xdr:wsDr>", "")); // truncated XML
        boolean failsClosed;
        try {
            failsClosed = !WorkbookComparator.compare(a, broken).identical();
        } catch (Exception e) {
            failsClosed = true; // batch mode reports this pair as ERROR
        }
        assertTrue(failsClosed, "a corrupt part must never compare identical");
    }

    @Test void detectsFooterTextChange() throws Exception {
        assertDetects("S1.footer.odd", wb -> wb.getSheetAt(0).getFooter().setCenter("Page &P of &N"));
    }

    @Test void detectsEvenPageHeader() throws Exception {
        assertDetects("S1.headerFooter", wb -> wb.getSheetAt(0).getEvenHeader().setCenter("Statement"));
    }

    @Test void detectsHiddenRow() throws Exception {
        assertDetects("S2.row[3]", wb -> wb.getSheetAt(1).getRow(2).setZeroHeight(true));
    }

    @Test void detectsNumberFormatChange() throws Exception {
        assertDetects("S2!C2", wb -> {
            XSSFCell c = wb.getSheetAt(1).getRow(1).getCell(2);
            XSSFCellStyle s = wb.createCellStyle();
            s.cloneStyleFrom(c.getCellStyle());
            s.setDataFormat(wb.createDataFormat().getFormat("0.00"));
            c.setCellStyle(s);
        });
    }

    @Test void detectsCommentChange() throws Exception {
        assertDetects("S1.comment[B3]", wb -> wb.getSheetAt(0).getRow(2).getCell(1)
            .getCellComment().setString(new XSSFRichTextString("Includes pending items.")));
    }

    @Test void detectsHyperlinkChange() throws Exception {
        assertDetects("S1.link[A3]", wb -> wb.getSheetAt(0).getRow(2).getCell(0)
            .getHyperlink().setAddress("https://example.com/help2"));
    }

    @Test void detectsTinyNumericDifference() throws Exception {
        assertDetects("S2!C3", wb -> wb.getSheetAt(1).getRow(2).getCell(2).setCellValue(-4.7500000001));
    }

    @Test void columnRangeSplitIsIrrelevant() throws Exception {
        Path a = build("a", false, wb -> {
            wb.getSheetAt(1).setColumnWidth(3, 5000);
            wb.getSheetAt(1).setColumnWidth(4, 5000);
        });
        Path b = build("b", false, wb -> {
            var col = wb.getSheetAt(1).getCTWorksheet().getColsArray(0).addNewCol();
            col.setMin(4); col.setMax(5); col.setWidth(5000 / 256.0); col.setCustomWidth(true);
        });
        assertIdentical(a, b);
    }

    @Test void cellWithoutStyleIndexInStyledColumnIsReported() throws Exception {
        // POI re-adds s= on write, so strip it from the saved XML, as other tools may emit it.
        Path a = build("a", false, wb -> {
            XSSFSheet sh = wb.getSheetAt(1);
            XSSFCellStyle colStyle = wb.createCellStyle();
            colStyle.setDataFormat(wb.createDataFormat().getFormat("0.0%"));
            sh.setDefaultColumnStyle(4, colStyle);
            sh.getRow(1).createCell(4).setCellValue(0.25);
        });
        Path b = dir.resolve("b.xlsx");
        rewriteEntry(a, b, "xl/worksheets/sheet2.xml",
            xml -> xml.replaceFirst("(<c r=\"E2\"[^>]*?) s=\"\\d+\"", "$1"));
        var r = WorkbookComparator.compare(a, b);
        assertTrue(r.differences().stream().anyMatch(d -> d.contains("S2!E2")), r.report(10));
    }

    private static void rewriteEntry(Path src, Path dst, String entry, UnaryOperator<String> edit) throws IOException {
        try (ZipFile in = new ZipFile(src.toFile()); ZipOutputStream out = new ZipOutputStream(Files.newOutputStream(dst))) {
            boolean edited = false;
            for (ZipEntry e : Collections.list(in.entries())) {
                byte[] data = in.getInputStream(e).readAllBytes();
                if (e.getName().equals(entry)) {
                    String before = new String(data, StandardCharsets.UTF_8), after = edit.apply(before);
                    edited = !after.equals(before);
                    data = after.getBytes(StandardCharsets.UTF_8);
                }
                out.putNextEntry(new ZipEntry(e.getName()));
                out.write(data);
                out.closeEntry();
            }
            assertTrue(edited, "fixture edit did not apply to " + entry);
        }
    }

    @Test void systemicMismatchListsEveryDifference() throws Exception {
        Path a = build("a", false, wb -> {}), b = build("b", false, wb -> {
            Sheet sh = wb.getSheetAt(1);
            for (int r = 5; r < 1200; r++) sh.createRow(r).createCell(0).setCellValue(r);
        });
        var r = WorkbookComparator.compare(a, b);
        assertEquals(1195, r.total(), "one difference per added cell, none dropped");
        assertTrue(r.report(Integer.MAX_VALUE).contains("S2!A1200"), "last difference is in the full report");
    }

    // ------------------------------------------------------------ coverage gaps (fail closed)

    /** Features the snapshot has no specific code for must still be compared or flagged. */
    @Test void unhandledFeaturesAreNeverSilentlyIgnored() throws Exception {
        Path base = build("base", false, wb -> {
            XSSFSheet sh = wb.getSheetAt(1);
            sh.groupRow(5, 7);
            sh.createRow(5).createCell(0).setCellValue(1);
            XSSFSimpleShape shape = wb.getSheetAt(0).getDrawingPatriarch()
                .createSimpleShape(new XSSFClientAnchor(0, 0, 0, 0, 8, 2, 10, 5));
            shape.setShapeType(ShapeTypes.RECT);
            shape.setText("Note");
            shape.setFillColor(200, 220, 255);
        });
        String sheet1 = "xl/worksheets/sheet1.xml", sheet2 = "xl/worksheets/sheet2.xml";
        Map<String, Object[]> cases = new LinkedHashMap<>();
        cases.put("error triangles suppressed", new Object[]{sheet2, (UnaryOperator<String>) x ->
            x.replace("<pageMargins", "<ignoredErrors><ignoredError sqref=\"A1\" numberStoredAsText=\"1\"/></ignoredErrors><pageMargins")});
        cases.put("summary rows above detail", new Object[]{sheet2, (UnaryOperator<String>) x ->
            x.replaceFirst("<dimension", "<sheetPr><outlinePr summaryBelow=\"0\"/></sheetPr><dimension")});
        cases.put("collapsed outline", new Object[]{sheet2, (UnaryOperator<String>) x ->
            x.replaceFirst("<row r=\"6\"", "<row r=\"6\" collapsed=\"1\"")});
        cases.put("full recalc on open", new Object[]{"xl/workbook.xml", (UnaryOperator<String>) x ->
            x.replace("</workbook>", "<calcPr fullCalcOnLoad=\"1\"/></workbook>")});
        cases.put("drawing objects hidden", new Object[]{"xl/workbook.xml", (UnaryOperator<String>) x ->
            x.replaceFirst("<workbookPr([^/]*)/>", "<workbookPr$1 showObjects=\"none\"/>")});
        cases.put("right-to-left text in a style", new Object[]{"xl/styles.xml", (UnaryOperator<String>) x ->
            x.replaceFirst("<xf ([^>]*)/>(\\s*</cellXfs>)", "<xf $1><alignment readingOrder=\"2\"/></xf>$2")});
        cases.put("font shadow", new Object[]{"xl/styles.xml", (UnaryOperator<String>) x ->
            x.replaceFirst("<font>", "<font><shadow/>")});
        cases.put("shape fill colour", new Object[]{"xl/drawings/drawing1.xml", (UnaryOperator<String>) x ->
            x.replaceFirst("(<a:srgbClr val=\")[0-9A-Fa-f]{6}", "$1FF0000")});
        cases.put("sheet background image", new Object[]{sheet1, (UnaryOperator<String>) x ->
            x.replace("<drawing ", "<picture r:id=\"rIdBg\"/><drawing ")});
        cases.put("extension block (e.g. sparklines)", new Object[]{sheet2, (UnaryOperator<String>) x ->
            x.replace("</worksheet>", "<extLst><ext uri=\"{05C60535-1F16-4fd2-B633-F4F36F0B64E0}\"/></extLst></worksheet>")});

        List<String> missed = new ArrayList<>();
        for (var c : cases.entrySet()) {
            Path variant = dir.resolve("variant.xlsx");
            @SuppressWarnings("unchecked") UnaryOperator<String> edit = (UnaryOperator<String>) c.getValue()[1];
            rewriteEntry(base, variant, (String) c.getValue()[0], edit);
            if (WorkbookComparator.compare(base, variant).identical()) missed.add(c.getKey());
        }
        assertTrue(missed.isEmpty(), "reported identical despite a real difference: " + missed);
    }

    /** Bookkeeping Excel adds on every save must not be reported as a difference. */
    @Test void excelHousekeepingIsNotADifference() throws Exception {
        Path a = build("a", false, wb -> {});
        Path step1 = dir.resolve("step1.xlsx"), step2 = dir.resolve("step2.xlsx"), b = dir.resolve("b.xlsx");
        rewriteEntry(a, step1, "xl/workbook.xml", x -> x
            .replaceFirst("<workbookPr", "<fileVersion appName=\"xl\" lastEdited=\"7\" rupBuild=\"27328\"/><workbookPr defaultThemeVersion=\"166925\"")
            .replaceFirst("<workbookView ", "<workbookView xWindow=\"0\" yWindow=\"0\" windowWidth=\"28800\" windowHeight=\"12300\" ")
            .replace("</workbook>", "<calcPr calcId=\"191029\"/></workbook>"));
        rewriteEntry(step1, step2, "xl/worksheets/sheet2.xml", x -> x
            .replaceFirst("<sheetFormatPr ", "<sheetFormatPr xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" x14ac:dyDescent=\"0.25\" ")
            .replaceFirst("</sheetView>", "<selection pane=\"bottomLeft\" sqref=\"A2\"/></sheetView>"));
        rewriteEntry(step2, b, "xl/worksheets/sheet1.xml", x -> x
            .replaceFirst("<pageMargins", "<sheetProtection algorithmName=\"SHA-512\" hashValue=\"abc\" saltValue=\"def\" spinCount=\"100000\" sheet=\"1\"/><pageMargins"));
        Path protectedA = dir.resolve("protectedA.xlsx");
        rewriteEntry(a, protectedA, "xl/worksheets/sheet1.xml", x -> x
            .replaceFirst("<pageMargins", "<sheetProtection algorithmName=\"SHA-512\" hashValue=\"xyz\" saltValue=\"uvw\" spinCount=\"100000\" sheet=\"true\"/><pageMargins"));
        // Same content, same protection, different salt/hash and Excel bookkeeping: identical.
        assertIdentical(protectedA, b);
    }

    @Test void streamingSxssfOutputMatchesXssf() throws Exception {
        Path a = dir.resolve("xssf.xlsx"), b = dir.resolve("sxssf.xlsx");
        try (Workbook wb = new XSSFWorkbook(); OutputStream o = Files.newOutputStream(a)) {
            fillPlain(wb); wb.write(o);
        }
        try (org.apache.poi.xssf.streaming.SXSSFWorkbook wb = new org.apache.poi.xssf.streaming.SXSSFWorkbook(50);
             OutputStream o = Files.newOutputStream(b)) {
            fillPlain(wb); wb.write(o);
        }
        assertIdentical(a, b); // shared strings vs inline strings must not matter
    }

    private static void fillPlain(Workbook wb) {
        CellStyle money = wb.createCellStyle();
        money.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
        Font bold = wb.createFont(); bold.setBold(true);
        CellStyle head = wb.createCellStyle(); head.setFont(bold); head.setBorderBottom(BorderStyle.THIN);
        for (String name : new String[]{"Summary", "Transactions"}) {
            Sheet s = wb.createSheet(name);
            s.setColumnWidth(1, 7000);
            Row h = s.createRow(0);
            for (int c = 0; c < 3; c++) { Cell x = h.createCell(c); x.setCellValue("Col " + c); x.setCellStyle(head); }
            for (int r = 1; r <= 200; r++) {
                Row row = s.createRow(r);
                row.createCell(0).setCellValue("Line " + r);
                row.createCell(1).setCellValue("Desc \"" + r + "\"\twith tab");
                Cell amt = row.createCell(2); amt.setCellValue(r * 1.25 - 100); amt.setCellStyle(money);
            }
            s.addMergedRegion(new CellRangeAddress(201, 201, 0, 2));
        }
    }

    // ------------------------------------------------------------ fixtures

    private void assertIdentical(Path a, Path b) throws IOException {
        var r = WorkbookComparator.compare(a, b);
        assertTrue(r.identical(), r.report(30));
    }

    private void assertDetects(String keyFragment, Consumer<XSSFWorkbook> mutation) throws Exception {
        var r = WorkbookComparator.compare(build("a", false, wb -> {}), build("b", false, mutation));
        assertFalse(r.identical(), "mutation went undetected");
        assertTrue(r.differences().stream().anyMatch(d -> d.contains(keyFragment)), r.report(30));
    }

    private Path build(String name, boolean reverseStyles, Consumer<XSSFWorkbook> customize) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            if (reverseStyles) wb.createCellStyle(); // shifts every style index
            XSSFFont bold = wb.createFont();
            bold.setBold(true);
            XSSFCellStyle money = wb.createCellStyle();
            money.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
            XSSFCellStyle header = wb.createCellStyle();
            header.setFont(bold);
            header.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0xD9, (byte) 0xE1, (byte) 0xF2}, null));
            header.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            header.setBorderBottom(BorderStyle.THIN);
            XSSFCellStyle date = wb.createCellStyle();
            date.setDataFormat(wb.createDataFormat().getFormat("mm/dd/yyyy"));

            XSSFSheet summary = wb.createSheet("Summary");
            XSSFSheet tx = wb.createSheet("Transactions");

            Cell title = summary.createRow(0).createCell(0);
            title.setCellValue("Account Statement");
            title.setCellStyle(header);
            summary.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
            Row r2 = summary.createRow(2);
            r2.createCell(0).setCellValue("Opening balance");
            Cell open = r2.createCell(1);
            open.setCellValue(1250.5);
            open.setCellStyle(money);
            Row r3 = summary.createRow(3);
            r3.createCell(0).setCellValue("Closing balance");
            Cell close = r3.createCell(1);
            close.setCellFormula("B3+SUM(Transactions!C2:C4)");
            close.setCellStyle(money);
            summary.setColumnWidth(0, 6000);

            XSSFHyperlink link = wb.getCreationHelper().createHyperlink(HyperlinkType.URL);
            link.setAddress("https://example.com/help");
            r2.getCell(0).setHyperlink(link);

            XSSFDrawing drawing = summary.createDrawingPatriarch();
            XSSFComment comment = drawing.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, 3, 2, 5, 5));
            comment.setString(new XSSFRichTextString("Includes pending items"));
            comment.setAuthor("ESU");
            open.setCellComment(comment);
            int pic = wb.addPicture(png(0x1F4E79), Workbook.PICTURE_TYPE_PNG);
            drawing.createPicture(new XSSFClientAnchor(0, 0, 0, 0, 5, 0, 7, 3), pic);

            String[] heads = {"Date", "Description", "Amount"};
            Row h = tx.createRow(0);
            for (int i = 0; i < heads.length; i++) {
                Cell c = h.createCell(i);
                c.setCellValue(heads[i]);
                c.setCellStyle(header);
            }
            Object[][] rows = {{45200.0, "Deposit", 500.0}, {45201.0, "Coffee", -4.75}, {45203.0, "Rent", -1200.0}};
            for (int i = 0; i < rows.length; i++) {
                Row r = tx.createRow(i + 1);
                Cell d = r.createCell(0);
                d.setCellValue((Double) rows[i][0]);
                d.setCellStyle(date);
                r.createCell(1).setCellValue((String) rows[i][1]);
                Cell amt = r.createCell(2);
                amt.setCellValue((Double) rows[i][2]);
                amt.setCellStyle(money);
            }
            tx.setColumnWidth(1, 8000);
            tx.createFreezePane(0, 1);

            wb.getCreationHelper().createFormulaEvaluator().evaluateAll();
            customize.accept(wb);

            Path f = dir.resolve(name + ".xlsx");
            try (OutputStream o = Files.newOutputStream(f)) { wb.write(o); }
            return f;
        }
    }

    private static byte[] png(int rgb) throws IOException {
        BufferedImage img = new BufferedImage(4, 4, BufferedImage.TYPE_INT_RGB);
        for (int x = 0; x < 4; x++) for (int y = 0; y < 4; y++) img.setRGB(x, y, rgb);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ImageIO.write(img, "png", out);
        return out.toByteArray();
    }
}
