package parity;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;
import java.util.function.UnaryOperator;
import java.util.zip.*;

import static org.junit.jupiter.api.Assertions.assertTrue;

/** A statement-like workbook exercising every feature the comparator inspects, plus zip editing. */
final class Fixtures {

    private Fixtures() {}

    static Path statement(Path file) throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFFont bold = wb.createFont();
            bold.setBold(true);
            XSSFCellStyle title = wb.createCellStyle();
            title.setFont(bold);
            XSSFCellStyle money = wb.createCellStyle();
            money.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
            money.setWrapText(true);

            XSSFSheet s1 = wb.createSheet("Statement");
            XSSFSheet s2 = wb.createSheet("Detail");

            Cell t = s1.createRow(0).createCell(0);
            t.setCellValue("Account Statement");
            t.setCellStyle(title);
            String[] labels = {"Opening", "Deposits", "Fees", "Closing"};
            for (int r = 1; r <= 4; r++) {
                Row row = s1.createRow(r);
                row.createCell(0).setCellValue(labels[r - 1]);
                Cell v = row.createCell(1);
                v.setCellValue(1250.5 * r);
                v.setCellStyle(money);
                row.createCell(3).setCellValue("Yes");
            }
            s1.getRow(1).createCell(2).setCellFormula("B2*2");
            s1.getRow(3).setHeightInPoints(30);
            s1.setColumnWidth(0, 5000);

            XSSFHyperlink link = wb.getCreationHelper().createHyperlink(HyperlinkType.URL);
            link.setAddress("https://example.com/help");
            s1.getRow(1).getCell(0).setHyperlink(link);

            XSSFDrawing drawing = s1.createDrawingPatriarch();
            XSSFComment comment = drawing.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, 2, 1, 4, 4));
            comment.setString(new XSSFRichTextString("Includes pending items"));
            comment.setAuthor("ESU");
            s1.getRow(1).getCell(1).setCellComment(comment);
            drawing.createPicture(new XSSFClientAnchor(0, 0, 0, 0, 5, 0, 7, 3), wb.addPicture(png(0x1F4E79), Workbook.PICTURE_TYPE_PNG));
            XSSFSimpleShape shape = drawing.createSimpleShape(new XSSFClientAnchor(0, 0, 0, 0, 5, 5, 8, 8));
            shape.setShapeType(ShapeTypes.RECT);
            shape.setText("Note");
            shape.setFillColor(200, 220, 255);

            XSSFSheetConditionalFormatting cf = s1.getSheetConditionalFormatting();
            XSSFConditionalFormattingRule rule = cf.createConditionalFormattingRule(ComparisonOperator.GT, "3000");
            rule.createFontFormatting().setFontColor(new XSSFColor(new byte[]{(byte) 0xFF, 0, 0}, null));
            cf.addConditionalFormatting(new CellRangeAddress[]{CellRangeAddress.valueOf("B2:B5")}, rule);

            DataValidationHelper dvh = s1.getDataValidationHelper();
            s1.addValidationData(dvh.createValidation(dvh.createExplicitListConstraint(new String[]{"Yes", "No"}),
                new CellRangeAddressList(1, 4, 3, 3)));

            for (int r = 9; r <= 12; r++) {
                Row row = s1.createRow(r);
                for (int c = 0; c < 3; c++) {
                    if (r == 9) row.createCell(c).setCellValue("Col" + c);
                    else row.createCell(c).setCellValue(r * 10 + c);
                }
            }
            XSSFTable table = s1.createTable(new AreaReference("A10:C13", SpreadsheetVersion.EXCEL2007));
            table.setName("Txns");
            table.setDisplayName("Txns");
            var info = table.getCTTable().addNewTableStyleInfo();
            info.setName("TableStyleMedium2");
            info.setShowRowStripes(true);

            s1.getHeader().setCenter("Statement");
            s1.setRowBreak(19);
            s1.setHorizontallyCenter(true);

            for (int r = 0; r < 6; r++) {
                Row row = s2.createRow(r);
                row.createCell(0).setCellValue("Line " + r);
                row.createCell(2).setCellValue(r * 1.5);
            }
            s2.createFreezePane(0, 1);
            var cf2 = s2.getSheetConditionalFormatting();
            var scale = cf2.createConditionalFormattingColorScaleRule();
            var colors = scale.getColorScaleFormatting();
            colors.setColors(new Color[]{new XSSFColor(new byte[]{(byte) 0xF8, 0x69, 0x6B}, null),
                new XSSFColor(new byte[]{(byte) 0xFF, (byte) 0xEB, (byte) 0x84}, null),
                new XSSFColor(new byte[]{0x63, (byte) 0xBE, 0x7B}, null)});
            cf2.addConditionalFormatting(new CellRangeAddress[]{CellRangeAddress.valueOf("C2:C6")}, scale);

            XSSFDrawing d2 = s2.createDrawingPatriarch();
            XSSFChart chart = d2.createChart(new XSSFClientAnchor(0, 0, 0, 0, 4, 0, 10, 12));
            chart.setTitleText("Balance");
            XDDFCategoryAxis x = chart.createCategoryAxis(AxisPosition.BOTTOM);
            XDDFValueAxis y = chart.createValueAxis(AxisPosition.LEFT);
            XDDFDataSource<String> cats = XDDFDataSourcesFactory.fromStringCellRange(s2, CellRangeAddress.valueOf("A2:A6"));
            XDDFNumericalDataSource<Double> vals = XDDFDataSourcesFactory.fromNumericCellRange(s2, CellRangeAddress.valueOf("C2:C6"));
            XDDFChartData data = chart.createData(ChartTypes.BAR, x, y);
            data.addSeries(cats, vals);
            ((XDDFBarChartData) data).setBarDirection(BarDirection.COL);
            chart.plot(data);

            var props = wb.getProperties();
            props.getCoreProperties().setTitle("Statement");
            props.getCustomProperties().addProperty("Classification", "Internal");

            wb.getCreationHelper().createFormulaEvaluator().evaluateAll();
            try (OutputStream o = Files.newOutputStream(file)) { wb.write(o); }
        }
        return file;
    }

    static byte[] png(int rgb) throws IOException {
        BufferedImage img = new BufferedImage(4, 4, BufferedImage.TYPE_INT_RGB);
        for (int x = 0; x < 4; x++) for (int y = 0; y < 4; y++) img.setRGB(x, y, rgb);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ImageIO.write(img, "png", out);
        return out.toByteArray();
    }

    /** Copies a package, applying text edits to some entries and adding new ones. Every edit must apply. */
    static Path rewrite(Path src, Path dst, Map<String, UnaryOperator<String>> edits, Map<String, byte[]> additions)
            throws IOException {
        Set<String> applied = new HashSet<>();
        try (ZipFile in = new ZipFile(src.toFile()); ZipOutputStream out = new ZipOutputStream(Files.newOutputStream(dst))) {
            for (ZipEntry e : Collections.list(in.entries())) {
                byte[] data = in.getInputStream(e).readAllBytes();
                UnaryOperator<String> edit = edits.get(e.getName());
                if (edit != null) {
                    String before = new String(data, StandardCharsets.UTF_8), after = edit.apply(before);
                    if (!after.equals(before)) applied.add(e.getName());
                    data = after.getBytes(StandardCharsets.UTF_8);
                }
                out.putNextEntry(new ZipEntry(e.getName()));
                out.write(data);
                out.closeEntry();
            }
            for (var add : additions.entrySet()) {
                out.putNextEntry(new ZipEntry(add.getKey()));
                out.write(add.getValue());
                out.closeEntry();
            }
        }
        assertTrue(applied.containsAll(edits.keySet()), "fixture edits did not apply to " + edits.keySet() + " (applied " + applied + ")");
        return dst;
    }

    static Path rewrite(Path src, Path dst, String entry, UnaryOperator<String> edit) throws IOException {
        return rewrite(src, dst, Map.of(entry, edit), Map.of());
    }

    static String entry(Path file, String name) throws IOException {
        try (ZipFile z = new ZipFile(file.toFile())) {
            ZipEntry e = z.getEntry(name);
            return e == null ? null : new String(z.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8);
        }
    }
}
