package parity;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ooxml.POIXMLTypeLoader;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.*;
import org.apache.xmlbeans.SchemaProperty;
import org.apache.xmlbeans.SchemaType;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTOfficeStyleSheet;
import org.openxmlformats.schemas.drawingml.x2006.main.ThemeDocument;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;

import javax.xml.namespace.QName;
import java.io.*;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.security.MessageDigest;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Canonical, rendering-level snapshot of an .xlsx workbook.
 *
 * <p>Every part of the package is compared: semantically where meaning matters (values, resolved
 * styles, layout), in canonical XML form for everything else, and by content hash for binary
 * parts. Keys starting with {@link #INFO} hold bookkeeping that does not change what a reader sees
 * or how the file behaves (application stamps, window geometry, salted password hashes, ...);
 * the comparator reports those separately.
 */
public final class WorkbookSnapshot {

    /** Key prefix for informational (bookkeeping) entries. */
    public static final String INFO = "info.";

    /** Candidate passwords used to verify salted protection hashes, which differ on every save. */
    public record Options(List<String> passwords) {
        public static final Options DEFAULT = new Options(List.of());

        public Options {
            passwords = List.copyOf(passwords);
        }
    }

    private static final int LAST_COL = SpreadsheetVersion.EXCEL2007.getLastColumnIndex();
    /** Canonical text longer than this is stored as a digest to bound memory. */
    private static final int MAX_INLINE = 4096;
    private static final String REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static final String OFFICE_NS = "urn:schemas-microsoft-com:office:office";

    // ---- worksheet
    /** Worksheet children compared by dedicated code; every other child is compared generically. */
    private static final Set<String> SHEET_HANDLED = Set.of("sheetPr", "dimension", "sheetViews", "sheetFormatPr",
        "cols", "sheetData", "sheetProtection", "autoFilter", "mergeCells", "conditionalFormatting",
        "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks",
        "colBreaks", "drawing", "legacyDrawing", "tableParts");
    private static final Set<String> SHEET_PR_INFO = Set.of("codeName", "filterMode"); // VBA name; derived
    private static final Set<String> SHEET_VIEW_INFO = Set.of("workbookViewId", "selection", "activePane"); // UI state
    private static final Set<String> SHEET_FORMAT_HANDLED = Set.of("defaultRowHeight", "defaultColWidth", "baseColWidth");
    private static final Set<String> SHEET_FORMAT_INFO = Set.of("outlineLevelRow", "outlineLevelCol"); // derived
    private static final Set<String> COL_HANDLED = Set.of("min", "max", "width", "style", "hidden", "customWidth",
        "outlineLevel");
    private static final Set<String> COL_INFO = Set.of("bestFit"); // an editing hint
    private static final Set<String> PAGE_SETUP_HANDLED = Set.of("paperSize", "scale", "firstPageNumber",
        "fitToWidth", "fitToHeight", "pageOrder", "orientation", "blackAndWhite", "draft", "useFirstPageNumber",
        "copies");
    private static final Set<String> PAGE_SETUP_INFO = Set.of("horizontalDpi", "verticalDpi"); // printer driver
    /** Salted hashes differ on every save: reported as info, verified via {@link Options#passwords}. */
    private static final Set<String> PROTECTION_HASHES = Set.of("password", "algorithmName", "hashValue",
        "saltValue", "spinCount", "workbookPassword", "workbookAlgorithmName", "workbookHashValue",
        "workbookSaltValue", "workbookSpinCount", "revisionsPassword", "revisionsAlgorithmName",
        "revisionsHashValue", "revisionsSaltValue", "revisionsSpinCount");

    // ---- workbook
    private static final Set<String> WORKBOOK_HANDLED = Set.of("workbookPr", "workbookProtection", "bookViews",
        "sheets", "definedNames", "calcPr", "fileSharing");
    /** Version stamps, recovery info and Excel's workbook-level extensions (calc features, original path). */
    private static final Set<String> WORKBOOK_INFO_CHILDREN = Set.of("fileVersion", "fileRecoveryPr", "extLst",
        "AlternateContent");
    private static final Set<String> WORKBOOK_PR_INFO = Set.of("codeName", "defaultThemeVersion", "filterPrivacy",
        "backupFile", "checkCompatibility", "autoCompressPictures", "promptedSolutions", "publishItems",
        "allowRefreshQuery", "hidePivotFieldList", "saveExternalLinkValues");
    private static final Set<String> WORKBOOK_VIEW_INFO = Set.of("xWindow", "yWindow", "windowWidth",
        "windowHeight", "tabRatio"); // application window geometry
    private static final Set<String> FILE_SHARING_INFO = Set.of("userName", "reservationPassword",
        "algorithmName", "hashValue", "saltValue", "spinCount");

    // ---- styles
    private static final Set<String> XF_HANDLED = Set.of("numFmtId", "fontId", "fillId", "borderId", "quotePrefix",
        "alignment", "protection");
    /** Linkage to named cell styles: shown in Excel's style gallery, not in the cell. */
    private static final Set<String> XF_INFO = Set.of("xfId", "applyNumberFormat", "applyFont", "applyFill",
        "applyBorder", "applyAlignment", "applyProtection");
    private static final Set<String> ALIGNMENT_HANDLED = Set.of("horizontal", "vertical", "textRotation",
        "wrapText", "indent", "shrinkToFit");
    private static final Set<String> FONT_HANDLED = Set.of("b", "i", "strike", "u", "vertAlign", "sz", "color",
        "name", "scheme");
    /** Only steer font substitution when the named font is missing. */
    private static final Set<String> FONT_INFO = Set.of("family", "charset");
    private static final Set<String> BORDER_HANDLED = Set.of("left", "right", "top", "bottom", "diagonal",
        "diagonalUp", "diagonalDown");
    private static final Set<String> STYLESHEET_HANDLED = Set.of("numFmts", "fonts", "fills", "borders", "cellXfs",
        "dxfs", "indexedColors");

    // ---- parts
    private static final String CORE_PROPS = "application/vnd.openxmlformats-package.core-properties+xml";
    private static final String APP_PROPS = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
    private static final String CUSTOM_PROPS = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
    /** Derived data or printer-driver blobs: compared, but as information. */
    private static final Set<String> INFO_PART_TYPES = Set.of(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings");

    private final Map<String, String> entries = new LinkedHashMap<>();
    private final List<String> uncovered = new ArrayList<>();
    private final Set<String> describedParts = new HashSet<>();
    private final String[] styleCache;                                // xf index -> style id
    private final Map<String, String> styleTable = new LinkedHashMap<>(); // style id -> description
    private final XSSFWorkbook wb;
    private final StylesTable styles;
    private final ThemesTable theme;
    private final SharedStringsTable sharedStrings;
    private final RichValue[] sharedStringCache;
    private final BitSet referencedStrings = new BitSet();
    private final Options options;
    private final String defaultStyle;

    private WorkbookSnapshot(XSSFWorkbook wb, Options options) {
        this.wb = wb;
        this.options = options;
        this.styles = wb.getStylesSource();
        this.theme = wb.getTheme();
        this.sharedStrings = wb.getSharedStringSource();
        this.sharedStringCache = new RichValue[sharedStrings == null ? 0 : sharedStrings.getUniqueCount()];
        this.styleCache = new String[styles.getNumCellStyles()];
        this.defaultStyle = styleAt(0);
        described(wb);
        described(styles);
        described(theme);
        described(sharedStrings);
        captureWorkbook();
        captureStylesheet();
        captureDocumentProperties();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) captureSheet(i);
        captureUnreferencedStrings();
        captureRemainingParts();
    }

    public static WorkbookSnapshot of(Path file) throws IOException {
        return of(file, Options.DEFAULT);
    }

    public static WorkbookSnapshot of(Path file, Options options) throws IOException {
        // Opening from a File gives POI random access to the zip (faster, less memory than a
        // stream). READ mode plus revert() guarantees the file on disk is never modified.
        OPCPackage pkg;
        try {
            pkg = OPCPackage.open(file.toFile(), PackageAccess.READ);
        } catch (InvalidFormatException e) {
            throw new IOException("Not a valid .xlsx: " + file, e);
        }
        try {
            return new WorkbookSnapshot(new XSSFWorkbook(pkg), options);
        } finally {
            pkg.revert();
        }
    }

    /** For generator output held in memory; snapshot the bytes customers would receive. */
    public static WorkbookSnapshot of(byte[] xlsx) throws IOException {
        return of(xlsx, Options.DEFAULT);
    }

    public static WorkbookSnapshot of(byte[] xlsx, Options options) throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
            return new WorkbookSnapshot(wb, options);
        }
    }

    public Map<String, String> entries() { return Collections.unmodifiableMap(entries); }
    public List<String> uncovered() { return Collections.unmodifiableList(uncovered); }
    /** Style id (content hash) -> full resolved style description. */
    public Map<String, String> styles() { return Collections.unmodifiableMap(styleTable); }

    /** Plain-text form, usable as a committed golden file. */
    public String toText() {
        StringBuilder sb = new StringBuilder();
        entries.forEach((k, v) -> sb.append(k).append(" = ").append(v).append('\n'));
        styleTable.forEach((k, v) -> sb.append("style ").append(k).append(" = ").append(v).append('\n'));
        uncovered.forEach(u -> sb.append("UNCOVERED ").append(u).append('\n'));
        return sb.toString();
    }

    // ---------------------------------------------------------------- workbook

    private void captureWorkbook() {
        put("wb.contentType", wb.getPackagePart().getContentType()); // macro-enabled or not
        put("wb.sheetCount", String.valueOf(wb.getNumberOfSheets()));
        put("wb.view", "active=" + wb.getActiveSheetIndex() + " firstVisibleTab=" + wb.getFirstVisibleTab());
        put("wb.date1904", String.valueOf(wb.isDate1904()));
        put("wb.protection", "structure=" + wb.isStructureLocked() + " windows=" + wb.isWindowsLocked());
        captureTheme();
        for (XSSFName n : wb.getAllNames()) {
            String scope = n.getSheetIndex() < 0 ? "wb" : "S" + (n.getSheetIndex() + 1);
            put("name[" + n.getNameName() + "@" + scope + "]", q(n.getRefersToFormula()));
        }
        CTWorkbook x = wb.getCTWorkbook();
        CTSheet[] sheets = x.getSheets().getSheetArray();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            put("S" + (i + 1) + ".sheet", "name=" + q(wb.getSheetName(i)) + " visibility=" + wb.getSheetVisibility(i));
            if (i < sheets.length) put(INFO + "S" + (i + 1) + ".sheetId", String.valueOf(sheets[i].getSheetId()));
        }
        split("wb.workbookPr", x.getWorkbookPr(), "workbookPr", Set.of("date1904"), WORKBOOK_PR_INFO, true);
        split("wb.calcPr", x.getCalcPr(), "calcPr", Set.of(), Set.of("calcId"), true); // calcId: Excel build
        if (x.isSetWorkbookProtection()) {
            CTWorkbookProtection p = x.getWorkbookProtection();
            split("wb.protectionXml", p, "workbookProtection", Set.of(), PROTECTION_HASHES, true);
            if (p.isSetWorkbookPassword() || p.isSetWorkbookHashValue()) verifyPassword("wb.password", wb::validateWorkbookPassword);
        }
        if (x.isSetFileSharing()) {
            split("wb.fileSharing", x.getFileSharing(), "fileSharing", Set.of(), FILE_SHARING_INFO, true);
        }
        if (x.isSetBookViews()) {
            CTBookView[] views = x.getBookViews().getWorkbookViewArray();
            for (int v = 0; v < views.length; v++) {
                split("wb.window" + v, views[v], "workbookView", Set.of(), WORKBOOK_VIEW_INFO, true);
            }
        }
        if (x.isSetDefinedNames()) {
            for (CTDefinedName n : x.getDefinedNames().getDefinedNameArray()) {
                String flags = canonical(n, "definedName", names(Set.of("name", "localSheetId", "#text")), true);
                if (!flags.equals("definedName")) {
                    put("nameFlags[" + n.getName() + "@" + (n.isSetLocalSheetId() ? "S" + (n.getLocalSheetId() + 1) : "wb") + "]", flags);
                }
            }
        }
        captureOtherChildren("wb", x, WORKBOOK_HANDLED, WORKBOOK_INFO_CHILDREN, wb.getPackagePart());
    }

    private void captureTheme() {
        if (theme == null) return;
        StringJoiner colors = new StringJoiner(",");
        for (int i = 0; i < 12; i++) {
            XSSFColor c = theme.getThemeColor(i);
            colors.add(c == null || c.getRGB() == null ? "-" : hex(c.getRGB()));
        }
        put("wb.themeColors", colors.toString());
        try (InputStream in = theme.getPackagePart().getInputStream()) {
            CTOfficeStyleSheet t = ThemeDocument.Factory.parse(in).getTheme();
            var fonts = t.getThemeElements().getFontScheme();
            // With font scheme MINOR/MAJOR, Excel renders the theme's fonts rather than the font name.
            put("wb.themeFonts", "major=" + q(fonts.getMajorFont().getLatin().getTypeface())
                + " minor=" + q(fonts.getMinorFont().getLatin().getTypeface()));
            put("wb.themeFontScheme", canonicalBounded(fonts, "fontScheme", names(Set.of("name"))));
            // Fills, lines and effects that shapes and charts reference by index.
            forEachChild(t.getThemeElements().getFmtScheme(), (name, child) ->
                put("wb.themeFormat." + name, canonicalBounded(child, name, NONE)));
            put(INFO + "wb.theme", canonicalBounded(t, "theme", names(Set.of("themeElements"))));
        } catch (Exception e) {
            uncovered.add("theme part unreadable: " + e.getMessage());
        }
    }

    private void captureStylesheet() {
        CTStylesheet ss = styles.getCTStylesheet();
        if (ss.isSetColors() && ss.getColors().isSetIndexedColors()) {
            // A custom palette changes what every indexed colour reference renders as.
            put("wb.indexedColors", canonical(ss.getColors().getIndexedColors(), "indexedColors", NONE, false));
        }
        // Named styles, the table-style gallery, recently used colours, style extensions.
        put(INFO + "styles.other", canonicalBounded(ss, "styleSheet", names(STYLESHEET_HANDLED)));
    }

    private void captureDocumentProperties() {
        POIXMLProperties props = wb.getProperties();
        var core = props.getCoreProperties();
        put(INFO + "docProps.core", "title=" + q(core.getTitle()) + " subject=" + q(core.getSubject())
            + " creator=" + q(core.getCreator()) + " keywords=" + q(core.getKeywords())
            + " description=" + q(core.getDescription()) + " category=" + q(core.getCategory())
            + " status=" + q(core.getContentStatus()) + " lastModifiedBy=" + q(core.getLastModifiedByUser())
            + " revision=" + q(core.getRevision()) + " created=" + core.getCreated() + " modified=" + core.getModified());
        put(INFO + "docProps.app", canonicalBounded(props.getExtendedProperties().getUnderlyingProperties(), "Properties", NONE));
        // Custom properties carry business metadata such as information-classification labels.
        for (var prop : props.getCustomProperties().getUnderlyingProperties().getPropertyArray()) {
            put("docProps.custom[" + prop.getName() + "]", canonical(prop, "property", names(Set.of("pid", "name")), false));
        }
        for (String type : List.of(CORE_PROPS, APP_PROPS, CUSTOM_PROPS)) {
            wb.getPackage().getPartsByContentType(type).forEach(p -> describedParts.add(p.getPartName().getName()));
        }
    }

    private void verifyPassword(String key, java.util.function.Predicate<String> validator) {
        List<String> candidates = options.passwords();
        if (candidates.isEmpty()) return; // the salted hash itself is reported as info
        String verdict = "matches none of the " + candidates.size() + " candidate passwords";
        for (int k = 0; k < candidates.size(); k++) {
            if (validator.test(candidates.get(k))) {
                verdict = "matches candidate #" + (k + 1);
                break;
            }
        }
        put(key, verdict);
    }

    // ---------------------------------------------------------------- sheet

    private void captureSheet(int index) {
        XSSFSheet sh = wb.getSheetAt(index);
        String p = "S" + (index + 1);
        described(sh);
        captureSheetSettings(sh, p);
        ColumnStyles cols = captureColumns(sh, p);
        captureRowsAndCells(sh, p, cols);
        for (CellRangeAddress m : sh.getMergedRegions()) put(p + ".merge[" + m.formatAsString() + "]", "merged");
        captureHyperlinks(sh, p);
        captureComments(sh, p);
        captureValidations(sh, p);
        captureConditionalFormatting(sh, p);
        captureTables(sh, p);
        captureDrawing(sh, p);
    }

    private void captureSheetSettings(XSSFSheet sh, String p) {
        CTWorksheet ws = sh.getCTWorksheet();
        // Canonical XML first: some POI getters below create missing elements as a side effect.
        split(p + ".sheetPr", ws.getSheetPr(), "sheetPr", Set.of("tabColor"), SHEET_PR_INFO, true);
        split(p + ".sheetFormatPr", ws.getSheetFormatPr(), "sheetFormatPr", SHEET_FORMAT_HANDLED, SHEET_FORMAT_INFO, true);
        if (ws.isSetDimension()) put(INFO + p + ".dimension", ws.getDimension().getRef());
        if (ws.isSetSheetViews()) {
            CTSheetView[] views = ws.getSheetViews().getSheetViewArray();
            for (int v = 0; v < views.length; v++) {
                split(p + ".sheetView" + v, views[v], "sheetView", Set.of(), SHEET_VIEW_INFO, true);
            }
        }
        if (ws.isSetSheetProtection()) {
            CTSheetProtection prot = ws.getSheetProtection();
            split(p + ".protectionXml", prot, "sheetProtection", Set.of(), PROTECTION_HASHES, true);
            if (prot.isSetPassword() || prot.isSetHashValue()) verifyPassword(p + ".password", sh::validateSheetPassword);
        }
        put(p + ".printOptionsXml", canonical(ws.getPrintOptions(), "printOptions", NONE, true));
        if (ws.isSetRowBreaks()) put(p + ".rowBreaks", canonical(ws.getRowBreaks(), "rowBreaks", names(Set.of("count", "manualBreakCount")), false));
        if (ws.isSetColBreaks()) put(p + ".colBreaks", canonical(ws.getColBreaks(), "colBreaks", names(Set.of("count", "manualBreakCount")), false));
        captureOtherChildren(p, ws, SHEET_HANDLED, Set.of(), sh.getPackagePart());

        String view = "gridlines=" + sh.isDisplayGridlines() + " headings=" + sh.isDisplayRowColHeadings()
            + " zeros=" + sh.isDisplayZeros() + " rtl=" + sh.isRightToLeft() + " formulas=" + sh.isDisplayFormulas()
            + " selected=" + sh.isSelected() + " activeCell=" + sh.getActiveCell();
        if (ws.isSetSheetViews() && ws.getSheetViews().sizeOfSheetViewArray() > 0) {
            CTSheetView v = ws.getSheetViews().getSheetViewArray(0);
            view += " zoom=" + (v.isSetZoomScale() ? v.getZoomScale() : 100)
                + " mode=" + (v.isSetView() ? v.getView() : "normal")
                + " topLeft=" + (v.isSetTopLeftCell() ? v.getTopLeftCell() : "A1");
        }
        put(p + ".view", view);

        PaneInformation pane = sh.getPaneInformation();
        if (pane != null) {
            put(p + ".pane", (pane.isFreezePane() ? "freeze" : "split")
                + " x=" + pane.getVerticalSplitPosition() + " y=" + pane.getHorizontalSplitPosition()
                + " leftCol=" + pane.getVerticalSplitLeftColumn() + " topRow=" + pane.getHorizontalSplitTopRow());
        }

        CTSheetFormatPr f = ws.getSheetFormatPr();
        put(p + ".defaults", "rowHeight=" + sh.getDefaultRowHeight()
            + " baseColWidth=" + sh.getDefaultColumnWidth()
            + " defaultColWidth=" + (f != null && f.isSetDefaultColWidth() ? f.getDefaultColWidth() : "-"));

        XSSFPrintSetup ps = sh.getPrintSetup();
        put(p + ".print", "paper=" + ps.getPaperSize() + " landscape=" + ps.getLandscape()
            + " scale=" + ps.getScale() + " fitToPage=" + sh.getFitToPage()
            + " fit=" + ps.getFitWidth() + "x" + ps.getFitHeight()
            + " pageStart=" + (ps.getUsePage() ? ps.getPageStart() : "auto")
            + " leftToRight=" + ps.getLeftToRight() + " bw=" + ps.getNoColor() + " draft=" + ps.getDraft()
            + " copies=" + ps.getCopies()
            + " center=" + sh.getHorizontallyCenter() + "/" + sh.getVerticallyCenter()
            + " printGridlines=" + sh.isPrintGridlines() + " printHeadings=" + sh.isPrintRowAndColumnHeadings()
            + " margins=" + margins(sh));
        split(p + ".pageSetupXml", ws.getPageSetup(), "pageSetup", PAGE_SETUP_HANDLED, PAGE_SETUP_INFO, true);

        // Raw header/footer XML: POI's getEvenHeader()/getFirstHeader() set the
        // differentOddEven/differentFirst flags as a side effect, which would mask real diffs.
        if (ws.isSetHeaderFooter()) {
            CTHeaderFooter hf = ws.getHeaderFooter();
            boolean defaults = !hf.getDifferentFirst() && !hf.getDifferentOddEven()
                && hf.getScaleWithDoc() && hf.getAlignWithMargins();
            if (!defaults) { // an empty <headerFooter/> renders exactly like none at all
                put(p + ".headerFooter", "differentFirst=" + hf.getDifferentFirst()
                    + " differentOddEven=" + hf.getDifferentOddEven()
                    + " scaleWithDoc=" + hf.getScaleWithDoc() + " alignWithMargins=" + hf.getAlignWithMargins());
            }
            headerFooter(p + ".header.odd", hf.getOddHeader());
            headerFooter(p + ".footer.odd", hf.getOddFooter());
            headerFooter(p + ".header.even", hf.getEvenHeader());
            headerFooter(p + ".footer.even", hf.getEvenFooter());
            headerFooter(p + ".header.first", hf.getFirstHeader());
            headerFooter(p + ".footer.first", hf.getFirstFooter());
        }

        if (sh.isSheetLocked()) {
            put(p + ".protection", "formatCells=" + sh.isFormatCellsLocked()
                + " formatCols=" + sh.isFormatColumnsLocked() + " formatRows=" + sh.isFormatRowsLocked()
                + " insertRows=" + sh.isInsertRowsLocked() + " deleteRows=" + sh.isDeleteRowsLocked()
                + " selectLocked=" + sh.isSelectLockedCellsLocked()
                + " selectUnlocked=" + sh.isSelectUnlockedCellsLocked()
                + " sort=" + sh.isSortLocked() + " autoFilter=" + sh.isAutoFilterLocked());
        }
        if (sh.getTabColor() != null) put(p + ".tabColor", color(sh.getTabColor()));
        if (ws.isSetAutoFilter()) put(p + ".autoFilter", canonical(ws.getAutoFilter(), "autoFilter", NONE, false));
    }

    @SuppressWarnings("deprecation") // short-constant overload compiles across POI 5.x
    private static String margins(XSSFSheet sh) {
        return sh.getMargin(Sheet.LeftMargin) + "," + sh.getMargin(Sheet.RightMargin) + ","
            + sh.getMargin(Sheet.TopMargin) + "," + sh.getMargin(Sheet.BottomMargin) + ","
            + sh.getMargin(Sheet.HeaderMargin) + "," + sh.getMargin(Sheet.FooterMargin);
    }

    /** Raw &L/&C/&R code string, so fonts and fields inside headers are compared too. */
    private void headerFooter(String key, String raw) {
        if (!isEmpty(raw)) put(key, q(raw));
    }

    // ---------------------------------------------------------------- columns

    /** Column style lookup over the sheet's elementary column intervals. */
    private record ColumnStyles(int[] from, String[] style) {
        String at(int col) {
            int i = Arrays.binarySearch(from, col);
            return style[i >= 0 ? i : -i - 2]; // from[0] == 0, so this never underflows
        }
    }

    /**
     * Emits canonical runs over all 16,384 columns, independent of how <cols> ranges are split.
     * Properties are constant between consecutive <col> boundaries, so they are read once per
     * interval rather than once per column (each POI column getter scans the whole <cols> list).
     */
    private ColumnStyles captureColumns(XSSFSheet sh, String p) {
        TreeSet<Integer> boundaries = new TreeSet<>();
        boundaries.add(0);
        for (CTCols cols : sh.getCTWorksheet().getColsArray()) {
            for (CTCol col : cols.getColArray()) {
                String range = "[" + col.getMin() + ":" + col.getMax() + "]";
                split(p + ".colXml" + range, col, "col", COL_HANDLED, COL_INFO, true, "col");
                long min = col.getMin() - 1, max = col.getMax() - 1; // 0-based, inclusive
                if (min >= 0 && min <= LAST_COL) boundaries.add((int) min);
                if (max >= 0 && max < LAST_COL) boundaries.add((int) max + 1);
            }
        }
        int[] from = boundaries.stream().mapToInt(Integer::intValue).toArray();
        String[] styleOf = new String[from.length];
        String prev = null;
        int runStart = 0;
        for (int i = 0; i < from.length; i++) {
            int c = from[i];
            styleOf[i] = columnStyle(sh, c);
            String props = columnProps(sh, c, styleOf[i]);
            if (!props.equals(prev)) {
                if (prev != null) putRun(p, runStart, c - 1, prev);
                prev = props;
                runStart = c;
            }
        }
        putRun(p, runStart, LAST_COL, prev);
        return new ColumnStyles(from, styleOf);
    }

    private String columnStyle(XSSFSheet sh, int c) {
        CellStyle cs = sh.getColumnStyle(c);
        return cs == null ? defaultStyle : style((XSSFCellStyle) cs);
    }

    private static String columnProps(XSSFSheet sh, int c, String styleFp) {
        return "w=" + sh.getColumnWidth(c)
            + (sh.isColumnHidden(c) ? " hidden" : "")
            + (sh.getColumnOutlineLevel(c) > 0 ? " outline=" + sh.getColumnOutlineLevel(c) : "")
            + " " + styleFp;
    }

    private void putRun(String p, int from, int to, String props) {
        put(p + ".cols[" + CellReference.convertNumToColString(from) + ":"
            + CellReference.convertNumToColString(to) + "]", props);
    }

    // ---------------------------------------------------------------- rows & cells

    private void captureRowsAndCells(XSSFSheet sh, String p, ColumnStyles cols) {
        short defaultHeight = sh.getDefaultRowHeight();
        for (Row row : sh) {
            XSSFRow r = (XSSFRow) row;
            CTRow ctRow = r.getCTRow();
            String rowStyle = r.isFormatted() ? styleAt((int) ctRow.getS()) : null;
            boolean collapsed = ctRow.getCollapsed(), thick = ctRow.getThickTop() || ctRow.getThickBot();
            boolean phonetic = ctRow.getPh();
            if (r.getHeight() != defaultHeight || r.getZeroHeight() || r.getOutlineLevel() > 0 || rowStyle != null
                    || collapsed || thick || phonetic) {
                put(p + ".row[" + (r.getRowNum() + 1) + "]", "h=" + r.getHeight()
                    + (r.getZeroHeight() ? " hidden" : "")
                    + (r.getOutlineLevel() > 0 ? " outline=" + r.getOutlineLevel() : "")
                    + (collapsed ? " collapsed" : "")
                    + (thick ? " thick=" + ctRow.getThickTop() + "/" + ctRow.getThickBot() : "")
                    + (phonetic ? " phonetic" : "")
                    + " " + (rowStyle == null ? "style:-" : rowStyle));
            }
            // customHeight=0 lets Excel re-fit the row after an edit; the stored height is compared above.
            if (ctRow.getCustomHeight()) put(INFO + p + ".rowCustomHeight[" + (r.getRowNum() + 1) + "]", "true");
            for (Cell cell : r) {
                XSSFCell c = (XSSFCell) cell;
                CTCell ct = c.getCTCell();
                String inherited = rowStyle != null ? rowStyle : cols.at(c.getColumnIndex());
                // Raw style index: POI's getCellStyle() allocates an object per cell. A missing s is
                // ambiguous when the row/column carries a style: the schema default is xf 0, but POI
                // and LibreOffice render the row/column style (verified in RenderEquivalenceTest).
                // Rather than pick a side, any encoding difference there is reported.
                // One XML read in the common case: presence only matters for s=0 under a styled row/column.
                long s = ct.getS();
                String fp = s != 0 || inherited.equals(defaultStyle) ? styleAt((int) s)
                    : ct.isSetS() ? defaultStyle : "noStyleIndex(inherits " + inherited + " or xf0)";
                CellType type = c.getCellType(); // each POI accessor is an XML read; fetch once
                if (type == CellType.BLANK && fp.equals(inherited)) continue;
                String key = p + "!" + c.getReference();
                StringBuilder info = new StringBuilder();
                StringBuilder v = new StringBuilder(value(c, ct, type, info));
                if (ct.getPh()) v.append(" phonetic");
                if (ct.isSetCm()) v.append(" cm=").append(ct.getCm()); // dynamic-array / cell metadata
                if (ct.isSetVm()) v.append(" vm=").append(ct.getVm()); // rich value (e.g. image in cell)
                if (ct.isSetExtLst()) v.append(" ext=").append(canonicalBounded(ct.getExtLst(), "extLst", NONE));
                if (type == CellType.FORMULA && ct.isSetF()) v.append(formulaExtras(ct.getF(), info));
                put(key, v.append(" | ").append(fp).toString());
                if (info.length() > 0) put(INFO + key, info.toString().trim());
            }
        }
    }

    private String value(XSSFCell c, CTCell ct, CellType type, StringBuilder info) {
        return switch (type) {
            case NUMERIC -> "N:" + number(ct); // exact, no epsilon
            case STRING -> "S:" + stringValue(c, ct, info);
            case BOOLEAN -> "B:" + c.getBooleanCellValue();
            case ERROR -> "E:" + error(c.getErrorCellValue());
            case BLANK -> "BLANK";
            case FORMULA -> "F:" + formula(c) + " => " + cached(c)
                + (c.isPartOfArrayFormulaGroup() ? " array=" + c.getArrayFormulaRange().formatAsString() : "");
            default -> "?" + type;
        };
    }

    /** The stored value parsed as POI does, without POI re-deriving the cell type first. */
    private static String number(CTCell ct) {
        String v = ct.getV();
        if (v == null || v.isEmpty()) return "0.0";
        try {
            return Double.toString(Double.parseDouble(v));
        } catch (NumberFormatException e) {
            return "<not a number: " + q(v) + ">";
        }
    }

    private static String formula(XSSFCell c) {
        try {
            return q(c.getCellFormula());
        } catch (RuntimeException e) { // e.g. what-if data tables; their XML is compared in formulaExtras
            return "<" + e.getClass().getSimpleName() + ">";
        }
    }

    /** Data-table definitions and the bx flag; shared-formula bookkeeping (t/ref/si) is encoding only. */
    private static String formulaExtras(CTCellFormula f, StringBuilder info) {
        boolean dataTable = f.getT() == STCellFormulaType.DATA_TABLE;
        Set<String> skip = dataTable ? Set.of("si", "ca", "aca", "#text") : Set.of("t", "ref", "si", "ca", "aca", "#text");
        if (f.getCa() || f.getAca()) info.append(" calcFlags=").append(f.getCa() ? "ca" : "").append(f.getAca() ? "aca" : "");
        String extra = canonical(f, "f", names(skip), false);
        return extra.equals("f") ? "" : " " + extra;
    }

    private String cached(XSSFCell c) {
        return switch (c.getCachedFormulaResultType()) {
            case NUMERIC -> "N:" + Double.toString(c.getNumericCellValue());
            case STRING -> "S:" + q(c.getStringCellValue());
            case BOOLEAN -> "B:" + c.getBooleanCellValue();
            case ERROR -> "E:" + error(c.getErrorCellValue());
            default -> "none";
        };
    }

    private static String error(byte code) {
        try { return FormulaError.forInt(code).getString(); }
        catch (IllegalArgumentException e) { return "#code" + code; }
    }

    /** Rendered text with formatting runs and phonetic guides; info collects run-font substitution hints. */
    private record RichValue(String value, String info) {}

    private String stringValue(XSSFCell c, CTCell ct, StringBuilder info) {
        RichValue rv;
        if (ct.getT() == STCellType.S && sharedStrings != null) {
            int idx;
            try { idx = Integer.parseInt(ct.getV()); } catch (RuntimeException e) { idx = -1; }
            if (idx >= 0 && idx < sharedStringCache.length) {
                referencedStrings.set(idx);
                rv = sharedStringCache[idx];
                if (rv == null) rv = sharedStringCache[idx] = rich((XSSFRichTextString) sharedStrings.getItemAt(idx));
            } else {
                rv = rich(c.getRichStringCellValue());
            }
        } else {
            rv = rich(c.getRichStringCellValue());
        }
        if (!rv.info().isEmpty()) info.append(' ').append(rv.info());
        return rv.value();
    }

    private RichValue rich(XSSFRichTextString s) {
        String text = q(s.getString());
        StringBuilder value = new StringBuilder(text), info = new StringBuilder();
        int runs = s.numFormattingRuns();
        if (runs > 0) {
            StringJoiner j = new StringJoiner("; ", " runs[", "]");
            for (int i = 0; i < runs; i++) {
                XSSFFont f = s.getFontOfFormattingRun(i);
                String extra = f == null ? "" : canonical(f.getCTFont(), "font", names(FONT_HANDLED, FONT_INFO), false);
                j.add(s.getIndexOfFormattingRun(i) + "+" + s.getLengthOfFormattingRun(i) + ":" + font(f)
                    + (extra.isEmpty() || extra.equals("font") ? "" : " " + extra));
                if (f != null) {
                    String hint = canonical(f.getCTFont(), "font", only("font", FONT_INFO), false);
                    if (!hint.equals("font")) info.append("run").append(i).append('=').append(hint).append(' ');
                }
            }
            value.append(j);
        }
        CTRst rst = s.getCTRst();
        if (rst.sizeOfRPhArray() > 0 || rst.isSetPhoneticPr()) {
            value.append(" phonetic=").append(canonical(rst, "si", names(Set.of("t", "r")), false));
        }
        return new RichValue(value.toString(), info.toString().trim());
    }

    /** Strings kept in the shared table but shown nowhere: invisible, yet still data in the file. */
    private void captureUnreferencedStrings() {
        if (sharedStrings == null) return;
        List<String> unused = new ArrayList<>();
        for (int i = referencedStrings.nextClearBit(0); i < sharedStringCache.length; i = referencedStrings.nextClearBit(i + 1)) {
            unused.add(sharedStrings.getItemAt(i).getString());
        }
        if (unused.isEmpty()) return;
        Collections.sort(unused);
        put(INFO + "sharedStrings.unreferenced", unused.size() + " string(s), sha256="
            + sha256(String.join("\u0000", unused).getBytes(StandardCharsets.UTF_8)).substring(0, 16));
    }

    // ---------------------------------------------------------------- styles

    private String style(XSSFCellStyle s) {
        return s == null ? defaultStyle : styleAt(s.getIndex());
    }

    /** By raw xf index, so per-cell lookups never allocate a POI XSSFCellStyle. */
    private String styleAt(int index) {
        if (index < 0 || index >= styleCache.length) return "style:invalid#" + index;
        String id = styleCache[index];
        if (id == null) {
            XSSFCellStyle s = styles.getStyleAt(index);
            String description = buildStyle(s);
            id = "st:" + sha256(description.getBytes(StandardCharsets.UTF_8)).substring(0, 12);
            styleTable.putIfAbsent(id, description);
            styleCache[index] = id;
            String info = styleInfo(s.getCoreXf());
            if (!info.isEmpty()) put(INFO + "style[" + id + "]", info);
        }
        return id;
    }

    private String buildStyle(XSSFCellStyle s) {
        CTXf xf = s.getCoreXf();
        // "; " separates properties so the comparator can report exactly which ones differ.
        return "fmt=" + q(s.getDataFormatString())
            + "; font{" + font(s.getFont()) + "}"
            + "; fill{" + fill(s) + "}"
            + "; border{" + border(s) + "}"
            + "; align=" + s.getAlignment() + "/" + s.getVerticalAlignment()
            + "; wrap=" + s.getWrapText() + "; indent=" + s.getIndention() + "; rot=" + s.getRotation()
            + "; shrink=" + s.getShrinkToFit() + "; quote=" + s.getQuotePrefixed()
            + "; prot=" + (s.getLocked() ? "L" : "-") + (s.getHidden() ? "H" : "-")
            + "; xfXml=" + canonical(xf, "xf", names(XF_HANDLED, XF_INFO), false)
            + "; alignXml=" + canonical(xf.getAlignment(), "alignment", names(ALIGNMENT_HANDLED), false)
            + "; fontXml=" + canonical(styles.getFontAt((int) xf.getFontId()).getCTFont(), "font", names(FONT_HANDLED, FONT_INFO), false)
            + "; borderXml=" + canonical(styles.getBorderAt((int) xf.getBorderId()).getCTBorder(), "border", names(BORDER_HANDLED), false);
    }

    /** Named-style linkage and font substitution hints: real XML, but invisible in the cell. */
    private String styleInfo(CTXf xf) {
        StringBuilder sb = new StringBuilder();
        String linkage = canonical(xf, "xf", only("xf", XF_INFO), false);
        if (!linkage.equals("xf")) sb.append(linkage);
        String hints = canonical(styles.getFontAt((int) xf.getFontId()).getCTFont(), "font", only("font", FONT_INFO), false);
        if (!hints.equals("font")) sb.append(' ').append(hints);
        CTStylesheet ss = styles.getCTStylesheet();
        if (xf.isSetXfId() && ss.isSetCellStyles()) {
            for (CTCellStyle named : ss.getCellStyles().getCellStyleArray()) {
                if (named.getXfId() == xf.getXfId()) sb.append(" namedStyle=").append(q(named.getName()));
            }
        }
        return sb.toString().trim();
    }

    private String font(XSSFFont f) {
        if (f == null) return "-";
        return q(f.getFontName()) + " " + f.getFontHeight() + "tw"
            + (f.getBold() ? " B" : "") + (f.getItalic() ? " I" : "") + (f.getStrikeout() ? " S" : "")
            + (f.getUnderline() != 0 ? " U" + f.getUnderline() : "")
            + (f.getTypeOffset() != 0 ? " off" + f.getTypeOffset() : "")
            + " scheme=" + f.getScheme() + " color=" + color(f.getXSSFColor());
    }

    private String fill(XSSFCellStyle s) {
        CTFill ct = styles.getFillAt((int) s.getCoreXf().getFillId()).getCTFill();
        if (ct.isSetGradientFill()) return "gradient " + canonical(ct.getGradientFill(), "gradientFill", NONE, false);
        FillPatternType pattern = s.getFillPattern();
        if (pattern == FillPatternType.NO_FILL) return "none";
        String fg = color(s.getFillForegroundColorColor());
        if (pattern == FillPatternType.SOLID_FOREGROUND) return "solid " + fg; // bg is invisible
        return pattern + " fg=" + fg + " bg=" + color(s.getFillBackgroundColorColor());
    }

    private String border(XSSFCellStyle s) {
        String out = side("T", s.getBorderTop(), s.getTopBorderXSSFColor())
            + side("B", s.getBorderBottom(), s.getBottomBorderXSSFColor())
            + side("L", s.getBorderLeft(), s.getLeftBorderXSSFColor())
            + side("R", s.getBorderRight(), s.getRightBorderXSSFColor());
        CTBorder ct = styles.getBorderAt((int) s.getCoreXf().getBorderId()).getCTBorder();
        if ((ct.getDiagonalUp() || ct.getDiagonalDown()) && ct.isSetDiagonal()) {
            out += " D" + (ct.getDiagonalUp() ? "/" : "") + (ct.getDiagonalDown() ? "\\" : "")
                + "=" + canonical(ct.getDiagonal(), "diagonal", NONE, false);
        }
        return out.isEmpty() ? "none" : out.trim();
    }

    private String side(String name, BorderStyle style, XSSFColor c) {
        return style == BorderStyle.NONE ? "" : " " + name + "=" + style + ":" + color(c);
    }

    /** Resolves theme, indexed and tinted colors to the RGB Excel actually renders. */
    private String color(XSSFColor c) {
        if (c == null) return "none";
        byte[] rgb = c.getRGB();
        if (rgb == null && c.isThemed() && theme != null) {
            XSSFColor base = theme.getThemeColor(c.getTheme());
            if (base != null) rgb = base.getRGB();
        }
        if (rgb != null) return hex(tint(rgb, c.getTint()));
        if (c.isAuto()) return "auto";
        if (c.isIndexed()) return "indexed:" + c.getIndex();
        if (c.isThemed()) return "theme:" + c.getTheme() + (c.getTint() != 0 ? "~" + c.getTint() : "") + "(no theme part)";
        return "unresolved";
    }

    /** Same tint algorithm Excel and POI use. */
    private static byte[] tint(byte[] rgb, double tint) {
        byte[] src = rgb.length == 4 ? Arrays.copyOfRange(rgb, 1, 4) : rgb;
        byte[] out = new byte[3];
        for (int i = 0; i < 3; i++) {
            int lum = src[i] & 0xFF;
            double v = tint > 0 ? lum * (1 - tint) + (255 - 255 * (1 - tint))
                     : tint < 0 ? lum * (1 + tint) : lum;
            out[i] = (byte) v;
        }
        return out;
    }

    /** Differential formats are referenced by index; compare what they contain instead. */
    private String dxf(long id) {
        CTStylesheet ss = styles.getCTStylesheet();
        if (!ss.isSetDxfs() || id < 0 || id >= ss.getDxfs().sizeOfDxfArray()) return "dxf#" + id + "(missing)";
        return canonical(styles.getDxfAt((int) id), "dxf", NONE, false);
    }

    /** Appends the content of every *dxfId reference found in an element. */
    private String resolveDxfs(XmlObject xml) {
        StringBuilder sb = new StringBuilder();
        visitAttributes(xml, "", (path, name, value) -> {
            String local = name.getLocalPart();
            if (name.getNamespaceURI().isEmpty() && (local.equals("dxfId") || local.endsWith("DxfId"))) {
                try { sb.append(' ').append(path).append('@').append(local).append('=').append(dxf(Long.parseLong(value))); }
                catch (NumberFormatException e) { sb.append(' ').append(path).append('@').append(local).append("=?").append(value); }
            }
        });
        return sb.toString();
    }

    // ---------------------------------------------------------------- annotations

    private void captureHyperlinks(XSSFSheet sh, String p) {
        for (XSSFHyperlink h : sh.getHyperlinkList()) {
            put(p + ".link[" + h.getCellRef() + "]", h.getType() + " " + q(h.getAddress())
                + " label=" + q(h.getLabel()) + " tip=" + q(h.getTooltip()));
        }
    }

    private void captureComments(XSSFSheet sh, String p) {
        Map<CellAddress, XSSFComment> all = sh.getCellComments();
        if (all.isEmpty()) return;
        for (POIXMLDocumentPart related : sh.getRelations()) {
            if (related instanceof Comments) described(related);
        }
        XSSFVMLDrawing vml = sh.getVMLDrawing(false);
        described(vml);
        for (Map.Entry<CellAddress, XSSFComment> e : new TreeMap<>(all).entrySet()) {
            XSSFComment c = e.getValue();
            ClientAnchor a = c.getClientAnchor();
            String box = a == null ? "-" : a.getCol1() + "," + a.getRow1() + "+" + a.getDx1() + "," + a.getDy1()
                + "->" + a.getCol2() + "," + a.getRow2() + "+" + a.getDx2() + "," + a.getDy2();
            com.microsoft.schemas.vml.CTShape shape = vml == null ? null : vml.findCommentShape(e.getKey().getRow(), e.getKey().getColumn());
            String fill = shape == null || shape.getFillcolor() == null ? "-" : shape.getFillcolor();
            RichValue text = rich(c.getString());
            String key = p + ".comment[" + e.getKey().formatAsString() + "]";
            put(key, "author=" + q(c.getAuthor()) + " visible=" + c.isVisible() + " box=" + box
                + " fill=" + fill + " text=" + text.value());
            if (!text.info().isEmpty()) put(INFO + key, text.info());
        }
    }

    private void captureValidations(XSSFSheet sh, String p) {
        CTWorksheet ws = sh.getCTWorksheet();
        if (!ws.isSetDataValidations()) return;
        CTDataValidations all = ws.getDataValidations();
        String container = canonical(all, "dataValidations", names(Set.of("dataValidation", "count")), true);
        if (!container.equals("dataValidations")) put(p + ".validations", container); // prompts disabled, ...
        for (CTDataValidation dv : all.getDataValidationArray()) {
            put(p + ".validation[" + sortedRefs(dv.xgetSqref().getStringValue()) + "]",
                canonical(dv, "dataValidation", names(Set.of("sqref")), false));
        }
    }

    private void captureConditionalFormatting(XSSFSheet sh, String p) {
        CTConditionalFormatting[] blocks = sh.getCTWorksheet().getConditionalFormattingArray();
        if (blocks.length == 0) return;
        // Relative order decides which rule wins; tools renumber the absolute priorities freely.
        List<Integer> order = Arrays.stream(blocks).flatMap(b -> Arrays.stream(b.getCfRuleArray()))
            .map(CTCfRule::getPriority).distinct().sorted().toList();
        for (CTConditionalFormatting block : blocks) {
            String ranges = sortedRefs(block.xgetSqref().getStringValue());
            String extra = canonical(block, "conditionalFormatting", names(Set.of("cfRule", "sqref")), true);
            CTCfRule[] rules = block.getCfRuleArray();
            for (int j = 0; j < rules.length; j++) {
                CTCfRule r = rules[j];
                put(p + ".cf[" + ranges + "]#" + j,
                    canonical(r, "cfRule", names(Set.of("dxfId", "priority")), false)
                        + " rank=" + order.indexOf(r.getPriority()) + resolveDxfs(r)
                        + (extra.equals("conditionalFormatting") ? "" : " " + extra));
            }
        }
    }

    private void captureTables(XSSFSheet sh, String p) {
        for (XSSFTable t : sh.getTables()) {
            described(t);
            CTTable ct = t.getCTTable();
            String styleName = ct.isSetTableStyleInfo() ? ct.getTableStyleInfo().getName() : null;
            put(p + ".table[" + ct.getName() + "]",
                canonical(ct, "table", (path, name) -> name.equals("id") || name.endsWith("DxfId"), false)
                    + resolveDxfs(ct) + customTableStyle(styleName));
        }
    }

    /** Custom table styles live in styles.xml and reference dxfs by index. */
    private String customTableStyle(String name) {
        CTStylesheet ss = styles.getCTStylesheet();
        if (name == null || !ss.isSetTableStyles()) return "";
        for (CTTableStyle ts : ss.getTableStyles().getTableStyleArray()) {
            if (name.equals(ts.getName())) {
                return " customStyle=" + canonical(ts, "tableStyle", names(Set.of("dxfId")), false) + resolveDxfs(ts);
            }
        }
        return "";
    }

    // ---------------------------------------------------------------- drawings

    /**
     * Every anchor in the drawing part: pictures, shapes, connectors, groups, charts, ink. Anchors
     * are keyed by position; their XML is compared canonically and every part they reference
     * (images, charts and the charts' own parts) is compared too.
     */
    private void captureDrawing(XSSFSheet sh, String p) {
        XSSFDrawing d = sh.getDrawingPatriarch();
        if (d == null) return;
        described(d);
        PackagePart source = d.getPackagePart();
        forEachChild(d.getCTDrawing(), (anchorName, anchor) -> {
            String kind = shapeKind(anchor);
            String key = p + "." + kind + "@" + anchorKey(anchorName, anchor);
            // Position is in the key; shape ids and names are generator bookkeeping. The shape's own
            // offset/size duplicate the anchor for top-level shapes (rotation and flips are kept);
            // inside groups they are real geometry and are compared.
            Skip skip = (path, name) ->
                (path.equals(anchorName) && Set.of("from", "to", "pos", "ext").contains(name))
                    || (path.endsWith("/cNvPr") && (name.equals("id") || name.equals("name")))
                    || ((name.equals("off") || name.equals("ext"))
                        && (path.equals(anchorName + "/sp/spPr/xfrm") || path.equals(anchorName + "/pic/spPr/xfrm")
                            || path.equals(anchorName + "/cxnSp/spPr/xfrm") || path.equals(anchorName + "/graphicFrame/xfrm")))
                    || (path.endsWith("/blip") && name.equals("cstate"));
            put(key, canonicalBounded(anchor, anchorName, skip));
            List<String> names = new ArrayList<>();
            visitAttributes(anchor, "", (path, name, value) -> {
                if (path.endsWith("/cNvPr") && name.getLocalPart().equals("name")) names.add(value);
            });
            if (!names.isEmpty()) put(INFO + key + ".names", String.join(",", names));
            describeReferences(key, anchor, source);
        });
    }

    private static String shapeKind(XmlObject anchor) {
        String[] kind = {"object"};
        forEachChild(anchor, (name, child) -> {
            switch (name) {
                case "sp" -> kind[0] = "shape";
                case "pic" -> kind[0] = "picture";
                case "cxnSp" -> kind[0] = "connector";
                case "grpSp" -> kind[0] = "group";
                case "graphicFrame" -> kind[0] = "frame";
                case "contentPart" -> kind[0] = "ink";
                case "AlternateContent" -> kind[0] = "alternate";
                default -> { }
            }
        });
        return kind[0];
    }

    private static String anchorKey(String anchorName, XmlObject anchor) {
        StringBuilder sb = new StringBuilder(switch (anchorName) {
            case "twoCellAnchor" -> "two:";
            case "oneCellAnchor" -> "one:";
            case "absoluteAnchor" -> "abs:";
            default -> anchorName + ":";
        });
        forEachChild(anchor, (name, child) -> {
            switch (name) {
                case "from" -> sb.append(marker(child));
                case "to" -> sb.append("->").append(marker(child));
                case "pos", "ext" -> sb.append(name).append(canonical(child, name, NONE, false).substring(name.length()));
                default -> { }
            }
        });
        return sb.toString();
    }

    private static String marker(XmlObject m) {
        Map<String, String> v = new HashMap<>();
        forEachChild(m, (name, child) -> {
            try (XmlCursor c = child.newCursor()) { v.put(name, c.getTextValue().strip()); }
        });
        return v.get("col") + "," + v.get("row") + "+" + v.get("colOff") + "," + v.get("rowOff");
    }

    // ---------------------------------------------------------------- referenced and remaining parts

    /** Unknown children are compared canonically, including any parts they reference. */
    private void captureOtherChildren(String p, XmlObject parent, Set<String> handled, Set<String> info,
                                      PackagePart source) {
        forEachChild(parent, (name, child) -> {
            if (handled.contains(name)) return;
            String key = (info.contains(name) ? INFO : "") + p + ".xml<" + name + ">";
            put(key, canonicalBounded(child, name, NONE));
            describeReferences(key, child, source);
        });
    }

    /** Resolves every relationship id inside {@code xml} and compares the parts they point to. */
    private void describeReferences(String key, XmlObject xml, PackagePart source) {
        List<String> ids = new ArrayList<>();
        visitAttributes(xml, "", (path, name, value) -> {
            if (isRelationshipId(name)) ids.add(value);
        });
        for (int i = 0; i < ids.size(); i++) {
            PackageRelationship rel = source.getRelationship(ids.get(i));
            String k = key + "->" + (i + 1);
            if (rel == null) put(k, "dangling relationship");
            else describeTarget(k, rel, source);
        }
    }

    private void describeTarget(String key, PackageRelationship rel, PackagePart source) {
        if (rel.getTargetMode() == TargetMode.EXTERNAL) {
            put(key, "external " + shortType(rel) + " " + rel.getTargetURI());
            return;
        }
        try {
            PackagePart target = source.getRelatedPart(rel);
            if (target == null) put(key, "missing part " + rel.getTargetURI());
            else describePart(key, target, true);
        } catch (InvalidFormatException e) {
            put(key, "unresolvable " + rel.getTargetURI());
        }
    }

    /**
     * Compares one part: XML parts canonically, split per top-level element so differences are
     * localized; binary parts by content hash. With {@code follow}, related parts are compared
     * too, in a deterministic order, so a chart's style, colours and embedded data are included.
     */
    private void describePart(String key, PackagePart part, boolean follow) {
        String name = part.getPartName().getName();
        if (!describedParts.add(name)) {
            put(key, "same part as described elsewhere: " + part.getContentType());
            return;
        }
        String type = part.getContentType();
        Set<String> referenced = new HashSet<>();
        XmlObject root = isXml(type) ? parse(part) : null;
        if (root != null) {
            put(key, "xml " + canonical(root, "root", (path, n) -> !path.contains("/") && !n.startsWith("@"), false)
                .replaceAll("\\{.*", "") + " type=" + type);
            forEachChild(root, (childName, child) -> put(key + "/" + childName, canonicalBounded(child, childName, NONE)));
            if (follow) {
                List<String> ids = new ArrayList<>();
                visitAttributes(root, "", (path, n, value) -> { if (isRelationshipId(n)) ids.add(value); });
                for (int i = 0; i < ids.size(); i++) {
                    PackageRelationship rel = part.getRelationship(ids.get(i));
                    referenced.add(ids.get(i));
                    if (rel != null) describeTarget(key + "->" + (i + 1), rel, part);
                }
            }
        } else {
            try (InputStream in = part.getInputStream()) {
                byte[] bytes = in.readAllBytes();
                put(key, "sha256=" + sha256(bytes) + " bytes=" + bytes.length + " type=" + type);
            } catch (IOException e) {
                uncovered.add("part " + name + " unreadable: " + e.getMessage());
            }
        }
        // Relationships not referenced from the XML (chart styles/colours) or external targets.
        try {
            List<PackageRelationship> rest = new ArrayList<>();
            for (PackageRelationship rel : part.getRelationships()) {
                if (!referenced.contains(rel.getId()) && (follow || rel.getTargetMode() == TargetMode.EXTERNAL)) rest.add(rel);
            }
            rest.sort(Comparator.comparing(PackageRelationship::getRelationshipType));
            for (int i = 0; i < rest.size(); i++) describeTarget(key + "~>" + shortType(rest.get(i)) + "#" + (i + 1), rest.get(i), part);
        } catch (InvalidFormatException e) {
            uncovered.add("relationships of " + name + " unreadable: " + e.getMessage());
        }
    }

    /** Everything nothing above reached: pivots, external links, VBA, metadata, custom XML, ... */
    private void captureRemainingParts() {
        List<PackagePart> parts;
        try {
            parts = new ArrayList<>(wb.getPackage().getParts());
        } catch (InvalidFormatException e) {
            uncovered.add("package unreadable: " + e.getMessage());
            return;
        }
        parts.sort(Comparator.comparing(part -> part.getPartName().getName()));
        for (PackagePart part : parts) {
            String name = part.getPartName().getName();
            if (part.isRelationshipPart() || describedParts.contains(name)) continue;
            String key = (INFO_PART_TYPES.contains(part.getContentType()) ? INFO : "") + "part[" + name + "]";
            describePart(key, part, false); // standalone, keyed by name: deterministic across files
        }
    }

    private static boolean isXml(String contentType) {
        return contentType.endsWith("xml") || contentType.endsWith("+xml") || contentType.endsWith("vmlDrawing");
    }

    /** Parses an XML part and returns its root element, or null when it is not well-formed XML. */
    private static XmlObject parse(PackagePart part) {
        try (InputStream in = part.getInputStream()) {
            XmlObject doc = XmlObject.Factory.parse(in, POIXMLTypeLoader.DEFAULT_XML_OPTIONS);
            try (XmlCursor c = doc.newCursor()) {
                return c.toFirstChild() ? c.getObject() : null;
            }
        } catch (Exception e) {
            return null; // e.g. legacy VML with unclosed tags: compared by content hash instead
        }
    }

    /** r:id/r:embed/... in OOXML; legacy VML (header/footer images) uses o:relid instead. */
    private static boolean isRelationshipId(QName name) {
        return REL_NS.equals(name.getNamespaceURI())
            || (OFFICE_NS.equals(name.getNamespaceURI()) && name.getLocalPart().equals("relid"));
    }

    private static String shortType(PackageRelationship rel) {
        String t = rel.getRelationshipType();
        return t.substring(t.lastIndexOf('/') + 1);
    }

    private void described(POIXMLDocumentPart part) {
        if (part != null) describedParts.add(part.getPackagePart().getPartName().getName());
    }

    // ---------------------------------------------------------------- canonical XML

    /** Decides, for an element at {@code path}, whether an attribute or child named {@code name} is left out. */
    @FunctionalInterface
    private interface Skip {
        boolean test(String path, String name);
    }

    private static final Skip NONE = (path, name) -> false;

    private static Skip names(Set<String> skip) {
        return (path, name) -> skip.contains(name);
    }

    private static Skip names(Set<String> a, Set<String> b) {
        return (path, name) -> a.contains(name) || b.contains(name);
    }

    /** Keeps only the given top-level attributes/children (everything beneath a kept child stays). */
    private static Skip only(String root, Set<String> keep) {
        return (path, name) -> path.equals(root) && !keep.contains(name);
    }

    /** Compares {@code handled} elsewhere, {@code info} names as bookkeeping, and the rest here. */
    private void split(String key, XmlObject xml, String name, Set<String> handled, Set<String> info, boolean dropEmpty) {
        split(key, xml, name, handled, info, dropEmpty, null);
    }

    /** With {@code omitIfBare}, entries equal to the bare element name are not recorded. */
    private void split(String key, XmlObject xml, String name, Set<String> handled, Set<String> info,
                       boolean dropEmpty, String omitIfBare) {
        String main = canonical(xml, name, names(handled, info), dropEmpty);
        if (omitIfBare == null || !main.equals(omitIfBare)) put(key, main);
        if (xml != null && !info.isEmpty()) {
            String extra = canonical(xml, name, only(name, info), false);
            if (!extra.equals(name)) put(INFO + key, extra);
        }
    }

    private static String canonicalBounded(XmlObject xml, String nameIfAbsent, Skip skip) {
        String s = canonical(xml, nameIfAbsent, skip, false);
        return s.length() <= MAX_INLINE ? s
            : "sha256:" + sha256(s.getBytes(StandardCharsets.UTF_8)) + " (" + s.length() + " chars canonical)";
    }

    /**
     * Order-independent, default-insensitive form of an XML element: attributes sorted, schema
     * defaults dropped, booleans and numbers normalized, namespaced attributes (relationship ids,
     * Excel revision ids and display hints) left out; relationship targets are compared separately
     * by {@link #describeReferences}. With {@code dropEmptyChildren}, a child that carries nothing
     * after normalization counts as absent; use it only for containers of settings blocks, never
     * where an element's mere presence means something.
     */
    private static String canonical(XmlObject xml, String nameIfAbsent, Skip skip, boolean dropEmptyChildren) {
        if (xml == null) return nameIfAbsent;
        StringBuilder sb = new StringBuilder();
        appendCanonical(xml, "", skip, dropEmptyChildren, sb, nameIfAbsent);
        return sb.toString();
    }

    private static String canonical(XmlObject xml, String nameIfAbsent, Set<String> skip, boolean dropEmptyChildren) {
        return canonical(xml, nameIfAbsent, names(skip), dropEmptyChildren);
    }

    /** {@code fragmentName} names a root that is a detached fragment (POI builds some fonts that way). */
    private static void appendCanonical(XmlObject xml, String parentPath, Skip skip, boolean dropEmpty, StringBuilder sb,
                                        String fragmentName) {
        SchemaType type = xml.schemaType();
        TreeMap<String, String> attrs = new TreeMap<>();
        List<XmlObject> children = new ArrayList<>();
        List<String> childNames = new ArrayList<>();
        String name, path, text = null;
        try (XmlCursor c = xml.newCursor()) {
            name = c.getName() == null ? fragmentName : c.getName().getLocalPart();
            path = parentPath.isEmpty() ? name : parentPath + "/" + name;
            if (c.toFirstAttribute()) {
                do {
                    QName qn = c.getName();
                    if (!qn.getNamespaceURI().isEmpty() || skip.test(path, qn.getLocalPart())) continue;
                    SchemaProperty prop = type.getAttributeProperty(qn);
                    String value = normalize(c.getTextValue(), prop);
                    if (prop != null && prop.getDefaultText() != null && normalize(prop.getDefaultText(), prop).equals(value)) continue;
                    attrs.put(qn.getLocalPart(), value);
                } while (c.toNextAttribute());
                c.toParent();
            }
            if (c.toFirstChild()) {
                do {
                    String child = c.getName().getLocalPart();
                    if (!skip.test(path, child)) {
                        children.add(c.getObject());
                        childNames.add(child);
                    }
                } while (c.toNextSibling());
            } else if (!skip.test(path, "#text")) {
                String t = c.getTextValue().strip();
                if (!t.isEmpty()) text = t;
            }
        }
        sb.append(name);
        if (!attrs.isEmpty()) sb.append(attrs);
        if (text != null) sb.append('=').append(q(text));
        List<String> kids = new ArrayList<>();
        for (int i = 0; i < children.size(); i++) {
            StringBuilder k = new StringBuilder();
            appendCanonical(children.get(i), path, skip, dropEmpty, k, childNames.get(i));
            if (!(dropEmpty && k.toString().equals(childNames.get(i)))) kids.add(k.toString());
        }
        if (!kids.isEmpty()) sb.append('{').append(String.join(";", kids)).append('}');
    }

    /** Makes "1"/"true", "0"/"false" and "15"/"15.0" compare equal where the schema type allows. */
    private static String normalize(String value, SchemaProperty prop) {
        if (prop == null || value == null) return value;
        SchemaType primitive = prop.getType().getPrimitiveType();
        if (primitive == null) return value;
        switch (primitive.getBuiltinTypeCode()) {
            case SchemaType.BTC_BOOLEAN:
                return value.equals("1") || value.equals("true") ? "true" : "false";
            case SchemaType.BTC_DECIMAL: case SchemaType.BTC_DOUBLE: case SchemaType.BTC_FLOAT:
                try { return new BigDecimal(value.strip()).stripTrailingZeros().toPlainString(); }
                catch (NumberFormatException e) { return value; }
            default:
                return value;
        }
    }

    @FunctionalInterface
    private interface ChildVisitor {
        void visit(String name, XmlObject child);
    }

    private static void forEachChild(XmlObject parent, ChildVisitor visitor) {
        if (parent == null) return;
        List<String> names = new ArrayList<>();
        List<XmlObject> kids = new ArrayList<>();
        try (XmlCursor c = parent.newCursor()) {
            if (c.toFirstChild()) {
                do {
                    names.add(c.getName().getLocalPart());
                    kids.add(c.getObject());
                } while (c.toNextSibling());
            }
        }
        for (int i = 0; i < kids.size(); i++) visitor.visit(names.get(i), kids.get(i));
    }

    @FunctionalInterface
    private interface AttributeVisitor {
        void visit(String path, QName name, String value);
    }

    /** Depth-first, document-order walk over every attribute in a subtree. */
    private static void visitAttributes(XmlObject xml, String parentPath, AttributeVisitor visitor) {
        List<XmlObject> kids = new ArrayList<>();
        String path;
        try (XmlCursor c = xml.newCursor()) {
            String name = c.getName() == null ? "#doc" : c.getName().getLocalPart();
            path = parentPath.isEmpty() ? name : parentPath + "/" + name;
            if (c.toFirstAttribute()) {
                do { visitor.visit(path, c.getName(), c.getTextValue()); } while (c.toNextAttribute());
                c.toParent();
            }
            if (c.toFirstChild()) {
                do { kids.add(c.getObject()); } while (c.toNextSibling());
            }
        }
        for (XmlObject k : kids) visitAttributes(k, path, visitor);
    }

    private static String sortedRefs(String sqref) {
        return Arrays.stream(sqref.trim().split("\\s+")).sorted().collect(Collectors.joining(","));
    }

    // ---------------------------------------------------------------- helpers

    private void put(String key, String value) {
        String k = key;
        int n = 2;
        while (entries.putIfAbsent(k, value) != null) k = key + "#" + n++;
    }

    private static boolean isEmpty(String s) { return s == null || s.isEmpty(); }

    private static String q(String s) {
        if (s == null) return "null";
        return '"' + s.replace("\\", "\\\\").replace("\n", "\\n").replace("\r", "\\r")
            .replace("\t", "\\t").replace("\"", "\\\"") + '"';
    }

    private static String hex(byte[] rgb) {
        byte[] b = rgb.length == 4 ? Arrays.copyOfRange(rgb, 1, 4) : rgb;
        return HexFormat.of().withUpperCase().formatHex(b);
    }

    private static String sha256(byte[] data) {
        try { return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(data)); }
        catch (Exception e) { throw new IllegalStateException(e); }
    }
}
