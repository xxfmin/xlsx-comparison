package parity;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.concurrent.TimeUnit;
import java.util.function.UnaryOperator;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

/**
 * Validates the comparator's rendering claims (G19) with a real renderer: every pair is printed to
 * PDF by LibreOffice, rasterized, and pixel-compared. The comparator's verdict must agree with
 * the render verdict. Controls with visible differences prove the method is sensitive; a file
 * against its own copy proves rendering is deterministic. LibreOffice stands in for Excel here.
 */
class RenderEquivalenceTest {

    @TempDir Path dir;

    /** {@code render}: true/false = what rendering must show; null = renderers disagree, not asserted. */
    record Pair(String name, Path a, Path b, boolean claimIdentical, Boolean render) {
        Pair(String name, Path a, Path b, boolean claimIdentical) { this(name, a, b, claimIdentical, claimIdentical); }
    }

    @Test
    void comparatorVerdictsMatchRenderedOutput() throws Exception {
        assumeTrue(onPath("soffice") && onPath("pdftoppm"), "LibreOffice and poppler-utils are needed");
        Path base = Fixtures.statement(dir.resolve("base.xlsx"));
        Path themed = edit(base, "themed", Map.of(), true);
        String s1 = "xl/worksheets/sheet1.xml";
        List<Pair> pairs = new ArrayList<>();

        // Controls
        pairs.add(new Pair("same file twice (determinism)", base, Files.copy(base, dir.resolve("copy.xlsx")), true));
        pairs.add(new Pair("CONTROL value changed", base, edit(base, "value", Map.of(s1, x -> x.replace("<v>2501.0</v>", "<v>2502.0</v>"))), false));
        pairs.add(new Pair("CONTROL highlight colour changed", base, edit(base, "cfcolor", Map.of("xl/styles.xml",
            x -> x.replace("<color rgb=\"FF0000\"/></font></dxf>", "<color rgb=\"00B050\"/></font></dxf>"))), false));
        pairs.add(new Pair("CONTROL row height changed", base, edit(base, "height", Map.of(s1, x -> x.replace("ht=\"30.0\"", "ht=\"45.0\""))), false));
        pairs.add(new Pair("CONTROL shape rotated", base, edit(base, "rot", Map.of("xl/drawings/drawing1.xml",
            x -> x.replaceFirst("(<xdr:sp>.*?<a:xfrm)>", "$1 rot=\"5400000\">"))), false));

        // Equivalences the comparator relies on
        pairs.add(new Pair("style table order", generated("plain", false, false), generated("reordered", true, false), true));
        pairs.add(new Pair("shared vs inline strings (SXSSF)", generated("xssf", false, false), generated("sxssf", false, true), true));
        pairs.add(new Pair("column range split", edit(base, "cols1", Map.of(s1, x -> x.replaceFirst("<cols>.*</cols>",
                "<cols><col min=\"1\" max=\"2\" width=\"19.53125\" customWidth=\"true\"/></cols>"))),
            edit(base, "cols2", Map.of(s1, x -> x.replaceFirst("<cols>.*</cols>",
                "<cols><col min=\"1\" max=\"1\" width=\"19.53125\" customWidth=\"true\"/><col min=\"2\" max=\"2\" width=\"19.53125\" customWidth=\"true\"/></cols>"))), true));
        pairs.add(new Pair("explicit blank vs absent cell", base, edit(base, "blank", Map.of(s1,
            x -> x.replace("<v>2</v></c></row><row r=\"3\">", "<v>2</v></c><c r=\"E2\" s=\"0\"/></row><row r=\"3\">"))), true));
        pairs.add(new Pair("theme colour vs identical RGB", edit(themed, "theme4", Map.of("xl/styles.xml",
                x -> x.replace("<b val=\"true\"/></font>", "<b val=\"true\"/><color theme=\"4\"/></font>"))),
            edit(themed, "rgb4", Map.of("xl/styles.xml",
                x -> x.replace("<b val=\"true\"/></font>", "<b val=\"true\"/><color rgb=\"FF4472C4\"/></font>"))), true));
        pairs.add(new Pair("explicit default attributes", base, edit(base, "defaults", Map.of(s1, x -> x
            .replace("<sheetView workbookViewId=\"0\" tabSelected=\"true\"/>", "<sheetView workbookViewId=\"0\" tabSelected=\"1\" showGridLines=\"1\" zoomScale=\"100\"/>")
            .replace("defaultRowHeight=\"15.0\"", "defaultRowHeight=\"15\"")
            .replace("<printOptions horizontalCentered=\"true\"/>", "<printOptions horizontalCentered=\"1\" gridLines=\"0\"/>"))), true));
        pairs.add(new Pair("conditional format priority renumbered", base, edit(base, "prio", Map.of(s1, x -> x.replace("priority=\"1\"", "priority=\"7\""))), true));
        pairs.add(new Pair("stale picture offset/size, same anchor", base, edit(base, "xfrm", Map.of("xl/drawings/drawing1.xml",
            x -> x.replace("<a:off x=\"0\" y=\"0\"/><a:ext cx=\"1066800\" cy=\"571500\"/>", "<a:off x=\"5000000\" y=\"900000\"/><a:ext cx=\"10\" cy=\"10\"/>"))), true));
        pairs.add(new Pair("shape ids renumbered", base, edit(base, "ids", Map.of("xl/drawings/drawing1.xml",
            x -> x.replace("<xdr:cNvPr id=\"2\" ", "<xdr:cNvPr id=\"42\" "))), true));

        // Ambiguous encoding: the schema default for a missing s is xf 0, but LibreOffice (and POI)
        // apply the column style. The comparator therefore reports it rather than guessing.
        String colStyled = "<cols><col min=\"1\" max=\"1\" width=\"19.53125\" customWidth=\"true\"/><col min=\"5\" max=\"5\" width=\"14\" style=\"1\" customWidth=\"1\"/></cols>";
        pairs.add(new Pair("cell without s in a styled column", edit(base, "withS", Map.of(s1, x -> x.replaceFirst("<cols>.*</cols>", colStyled)
                .replace("<v>2</v></c></row><row r=\"3\">", "<v>2</v></c><c r=\"E2\" s=\"1\" t=\"inlineStr\"><is><t>WWWW</t></is></c></row><row r=\"3\">"))),
            edit(base, "withoutS", Map.of(s1, x -> x.replaceFirst("<cols>.*</cols>", colStyled)
                .replace("<v>2</v></c></row><row r=\"3\">", "<v>2</v></c><c r=\"E2\" t=\"inlineStr\"><is><t>WWWW</t></is></c></row><row r=\"3\">"))), false, null));

        Map<Path, List<BufferedImage>> rendered = render(pairs.stream().flatMap(p -> Stream.of(p.a(), p.b())).distinct().toList());
        List<String> problems = new ArrayList<>();
        StringBuilder table = new StringBuilder();
        for (Pair p : pairs) {
            boolean comparatorSame = WorkbookComparator.compare(p.a(), p.b()).identical();
            boolean renderSame = samePixels(rendered.get(p.a()), rendered.get(p.b()));
            table.append(String.format("%-46s comparator=%-9s render=%-9s%n", p.name(),
                comparatorSame ? "identical" : "differs", renderSame ? "identical" : "differs"));
            if (comparatorSame != p.claimIdentical()) problems.add(p.name() + ": comparator verdict differs from the claim");
            if (p.render() != null && renderSame != p.render()) problems.add(p.name() + ": rendered output contradicts the claim");
        }
        System.out.println(table);
        assertTrue(problems.isEmpty(), String.join("\n", problems) + "\n" + table);
    }

    // ------------------------------------------------------------ fixtures

    private Path edit(Path src, String name, Map<String, UnaryOperator<String>> edits) throws Exception {
        return edit(src, name, edits, false);
    }

    private Path edit(Path src, String name, Map<String, UnaryOperator<String>> edits, boolean addTheme) throws Exception {
        Map<String, UnaryOperator<String>> all = new LinkedHashMap<>(edits);
        Map<String, byte[]> adds = new LinkedHashMap<>();
        if (addTheme) {
            adds.put("xl/theme/theme1.xml", GapCoverageTest.THEME.getBytes(java.nio.charset.StandardCharsets.UTF_8));
            all.put("[Content_Types].xml", x -> x.replace("</Types>",
                "<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/></Types>"));
            all.put("xl/_rels/workbook.xml.rels", x -> x.replace("</Relationships>",
                "<Relationship Id=\"rIdTheme\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/></Relationships>"));
        }
        return Fixtures.rewrite(src, dir.resolve(name + ".xlsx"), all, adds);
    }

    /** Same statement through different code paths: style creation order and streaming writer. */
    private Path generated(String name, boolean reverseStyles, boolean streaming) throws IOException {
        Path f = dir.resolve(name + ".xlsx");
        try (Workbook wb = streaming ? new SXSSFWorkbook(20) : new XSSFWorkbook(); OutputStream out = Files.newOutputStream(f)) {
            if (reverseStyles) wb.createCellStyle(); // shifts every style index
            Font bold = wb.createFont();
            bold.setBold(true);
            CellStyle head = wb.createCellStyle();
            head.setFont(bold);
            head.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
            head.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            head.setBorderBottom(BorderStyle.THIN);
            CellStyle money = wb.createCellStyle();
            money.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
            Sheet s = wb.createSheet("Statement");
            s.setColumnWidth(1, 7000);
            Row h = s.createRow(0);
            for (int c = 0; c < 3; c++) { Cell x = h.createCell(c); x.setCellValue("Head " + c); x.setCellStyle(head); }
            for (int r = 1; r <= 40; r++) {
                Row row = s.createRow(r);
                row.createCell(0).setCellValue("Line " + r);
                row.createCell(1).setCellValue("Merchant " + (r % 7));
                Cell amt = row.createCell(2);
                amt.setCellValue(r * 12.5 - 100);
                amt.setCellStyle(money);
            }
            s.addMergedRegion(new CellRangeAddress(41, 41, 0, 2));
            wb.write(out);
        }
        return f;
    }

    // ------------------------------------------------------------ rendering

    private Map<Path, List<BufferedImage>> render(List<Path> files) throws Exception {
        Path pdfDir = Files.createDirectories(dir.resolve("pdf"));
        Map<Path, List<BufferedImage>> out = new HashMap<>();
        // One LibreOffice process per batch, with a private profile so parallel runs don't collide.
        List<String> cmd = new ArrayList<>(List.of("soffice", "-env:UserInstallation=" + dir.resolve("lo-profile").toUri(),
            "--headless", "--convert-to", "pdf", "--outdir", pdfDir.toString()));
        Map<String, Path> byPdf = new HashMap<>();
        int n = 0;
        for (Path f : files) {
            // Repack into a standard zip container: LibreOffice rejects POI's streaming (SXSSF) container
            // (data descriptors, DOS date 1980-00-00), though Excel accepts it. Entry bytes are unchanged.
            Path unique = Fixtures.rewrite(f, dir.resolve("r" + (n++) + ".xlsx"), Map.of(), Map.of());
            cmd.add(unique.toString());
            byPdf.put(unique.getFileName().toString().replace(".xlsx", ".pdf"), f);
        }
        run(cmd, 600);
        for (var e : byPdf.entrySet()) {
            Path pdf = pdfDir.resolve(e.getKey());
            assertTrue(Files.exists(pdf), "LibreOffice did not produce " + pdf);
            Path prefix = pdfDir.resolve(e.getKey() + "-p");
            run(List.of("pdftoppm", "-r", "60", "-png", pdf.toString(), prefix.toString()), 120);
            List<BufferedImage> pages = new ArrayList<>();
            try (Stream<Path> s = Files.list(pdfDir)) {
                for (Path png : s.filter(x -> x.getFileName().toString().startsWith(e.getKey() + "-p")).sorted().toList()) {
                    pages.add(ImageIO.read(png.toFile()));
                }
            }
            out.put(e.getValue(), pages);
        }
        return out;
    }

    private static boolean samePixels(List<BufferedImage> a, List<BufferedImage> b) {
        if (a.size() != b.size()) return false;
        for (int i = 0; i < a.size(); i++) {
            BufferedImage x = a.get(i), y = b.get(i);
            if (x.getWidth() != y.getWidth() || x.getHeight() != y.getHeight()) return false;
            for (int px = 0; px < x.getWidth(); px++) {
                for (int py = 0; py < x.getHeight(); py++) {
                    if (x.getRGB(px, py) != y.getRGB(px, py)) return false;
                }
            }
        }
        return true;
    }

    private static void run(List<String> cmd, int timeoutSeconds) throws Exception {
        Process p = new ProcessBuilder(cmd).redirectErrorStream(true).redirectOutput(ProcessBuilder.Redirect.DISCARD).start();
        assertTrue(p.waitFor(timeoutSeconds, TimeUnit.SECONDS), "timed out: " + cmd.get(0));
    }

    private static boolean onPath(String tool) {
        return Stream.of(System.getenv("PATH").split(File.pathSeparator)).anyMatch(d -> Files.isExecutable(Path.of(d, tool)));
    }
}
