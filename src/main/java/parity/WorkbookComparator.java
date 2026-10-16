package parity;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.*;
import java.util.*;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.CompletionException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;

public final class WorkbookComparator {

    /**
     * Every difference found, in full. {@code differences} change what a reader sees or how the
     * file behaves; {@code info} are bookkeeping differences (application stamps, window geometry,
     * salted hashes, ...), which fail a comparison only in strict mode.
     */
    public record Result(String label, List<String> differences, List<String> info) {
        public Result(String label, List<String> differences) { this(label, differences, List.of()); }

        public boolean identical() { return differences.isEmpty(); }

        public boolean identicalStrict() { return differences.isEmpty() && info.isEmpty(); }

        public boolean passes(boolean strict) { return strict ? identicalStrict() : identical(); }

        public int total() { return differences.size(); }

        /** Console-friendly view; use {@code report(Integer.MAX_VALUE)} for everything. */
        public String report(int max) {
            StringBuilder sb = new StringBuilder(label).append(": ")
                .append(identical() ? "IDENTICAL" : total() + " difference(s)");
            if (!info.isEmpty()) sb.append(", ").append(info.size()).append(" informational");
            sb.append('\n');
            append(sb, differences, max);
            if (!info.isEmpty()) {
                sb.append("  Informational (bookkeeping only; fails in --strict mode):\n");
                append(sb, info, max);
            }
            return sb.toString();
        }

        private static void append(StringBuilder sb, List<String> items, int max) {
            items.stream().limit(max).forEach(d -> sb.append("  ").append(d).append('\n'));
            if (items.size() > max) sb.append("  ... and ").append(items.size() - max).append(" more\n");
        }
    }

    /**
     * Snapshots both workbooks concurrently, then compares. Batch callers that already run pairs
     * in parallel should use {@link #compareSequentially} to avoid nested parallelism.
     */
    public static Result compare(Path expected, Path actual) throws IOException {
        return compare(expected, actual, WorkbookSnapshot.Options.DEFAULT);
    }

    public static Result compare(Path expected, Path actual, WorkbookSnapshot.Options options) throws IOException {
        CompletableFuture<WorkbookSnapshot> other = CompletableFuture.supplyAsync(() -> {
            try { return WorkbookSnapshot.of(actual, options); }
            catch (IOException e) { throw new UncheckedIOException(e); }
        });
        WorkbookSnapshot exp = WorkbookSnapshot.of(expected, options);
        try {
            return compare(expected.getFileName().toString(), exp, other.join());
        } catch (CompletionException e) {
            if (e.getCause() instanceof UncheckedIOException u) throw u.getCause();
            throw e;
        }
    }

    public static Result compareSequentially(Path expected, Path actual) throws IOException {
        return compareSequentially(expected, actual, WorkbookSnapshot.Options.DEFAULT);
    }

    public static Result compareSequentially(Path expected, Path actual, WorkbookSnapshot.Options options)
            throws IOException {
        return compare(expected.getFileName().toString(),
            WorkbookSnapshot.of(expected, options), WorkbookSnapshot.of(actual, options));
    }

    public static Result compare(String label, WorkbookSnapshot expected, WorkbookSnapshot actual) {
        List<String> diffs = new ArrayList<>(), info = new ArrayList<>();
        // Fail closed: anything that could not be read blocks a pass, on either side.
        expected.uncovered().forEach(u -> diffs.add("UNREADABLE (expected) " + u));
        actual.uncovered().forEach(u -> diffs.add("UNREADABLE (actual)   " + u));

        Map<String, String> e = expected.entries(), a = actual.entries();
        e.forEach((key, ev) -> {
            String av = a.get(key);
            List<String> sink = key.startsWith(WorkbookSnapshot.INFO) ? info : diffs;
            if (av == null) sink.add("MISSING " + key + "\n      expected: " + ev);
            else if (!av.equals(ev)) sink.add("CHANGED " + key + "\n      expected: " + ev + "\n      actual:   " + av
                + styleDetail(ev, av, expected.styles(), actual.styles()));
        });
        a.forEach((key, av) -> {
            if (!e.containsKey(key)) {
                (key.startsWith(WorkbookSnapshot.INFO) ? info : diffs).add("EXTRA   " + key + "\n      actual:   " + av);
            }
        });
        return new Result(label, diffs, info);
    }

    private static final Pattern STYLE_ID = Pattern.compile("st:[0-9a-f]{12}");

    /** Explains a style-id mismatch as the specific style properties that differ. */
    private static String styleDetail(String ev, String av, Map<String, String> es, Map<String, String> as) {
        Matcher me = STYLE_ID.matcher(ev), ma = STYLE_ID.matcher(av);
        if (!me.find() || !ma.find() || me.group().equals(ma.group())) return "";
        List<String> e = List.of(es.getOrDefault(me.group(), "?").split("; "));
        List<String> a = List.of(as.getOrDefault(ma.group(), "?").split("; "));
        StringBuilder sb = new StringBuilder("\n      style differs:");
        for (int i = 0; i < Math.max(e.size(), a.size()); i++) {
            String x = i < e.size() ? e.get(i) : "-", y = i < a.size() ? a.get(i) : "-";
            if (!x.equals(y)) sb.append("\n        expected ").append(x).append("\n        actual   ").append(y);
        }
        return sb.toString();
    }

    /** One line per pair kept in memory; full detail goes straight to a report file. */
    private record Summary(String label, boolean passed, int differences, int info, String headline, Path report) {}

    /**
     * Batch mode:
     * java parity.WorkbookComparator [--strict] [--password=SECRET ...] expectedDir actualDir [reportDir]
     * Writes every difference for each pair that is not identical to reportDir/<file>.txt.
     */
    public static void main(String[] args) throws IOException {
        boolean strict = false;
        List<String> passwords = new ArrayList<>(), positional = new ArrayList<>();
        for (String arg : args) {
            if (arg.equals("--strict")) strict = true;
            else if (arg.startsWith("--password=")) passwords.add(arg.substring("--password=".length()));
            else positional.add(arg);
        }
        if (positional.size() < 2 || positional.size() > 3) {
            System.err.println("usage: WorkbookComparator [--strict] [--password=SECRET ...] <expectedDir> <actualDir> [reportDir]");
            System.exit(2);
        }
        Path expDir = Path.of(positional.get(0)), actDir = Path.of(positional.get(1));
        Path reportDir = Path.of(positional.size() == 3 ? positional.get(2) : "parity-reports");
        WorkbookSnapshot.Options options = new WorkbookSnapshot.Options(passwords);
        List<Path> files = xlsxIn(expDir);
        if (files.isEmpty()) {
            System.err.println("No .xlsx files in " + expDir + "; refusing to report success.");
            System.exit(2);
        }
        Files.createDirectories(reportDir);
        boolean strictMode = strict;
        List<Summary> summaries = files.parallelStream()
            .map(f -> run(f, actDir.resolve(f.getFileName()), reportDir, options, strictMode)).toList();

        Set<String> expectedNames = new HashSet<>();
        files.forEach(f -> expectedNames.add(f.getFileName().toString()));
        List<Path> orphans = xlsxIn(actDir).stream()
            .filter(f -> !expectedNames.contains(f.getFileName().toString())).toList();

        for (Summary s : summaries) {
            if (s.report() == null) continue;
            System.out.printf("%s %s: %d difference(s), %d informational, first: %s%n    full report: %s%n",
                s.passed() ? "PASS" : "FAIL", s.label(), s.differences(), s.info(), s.headline(), s.report());
        }
        orphans.forEach(o -> System.out.println("EXTRA FILE (no expected counterpart): " + o.getFileName()));
        long failed = summaries.stream().filter(s -> !s.passed()).count();
        System.out.printf("%d/%d workbooks pass%s%n", summaries.size() - failed, summaries.size(), strict ? " (strict)" : "");
        System.exit(failed == 0 && orphans.isEmpty() ? 0 : 1);
    }

    private static Summary run(Path expected, Path actual, Path reportDir, WorkbookSnapshot.Options options,
                               boolean strict) {
        String label = expected.getFileName().toString();
        Result result;
        if (!Files.exists(actual)) {
            result = new Result(label, List.of("candidate file missing: " + actual));
        } else {
            try {
                result = compareSequentially(expected, actual, options); // pairs already run in parallel
            } catch (Exception e) {
                result = new Result(label, List.of("ERROR " + e));
            }
        }
        boolean passed = result.passes(strict);
        if (result.identicalStrict()) return new Summary(label, true, 0, 0, "", null);
        Path report = reportDir.resolve(label + ".txt");
        try {
            Files.writeString(report, result.report(Integer.MAX_VALUE));
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
        List<String> first = result.differences().isEmpty() ? result.info() : result.differences();
        return new Summary(label, passed, result.total(), result.info().size(),
            first.get(0).lines().findFirst().orElse(""), report);
    }

    private static List<Path> xlsxIn(Path dir) throws IOException {
        if (!Files.isDirectory(dir)) return List.of();
        try (Stream<Path> s = Files.list(dir)) {
            return s.filter(f -> f.toString().toLowerCase().endsWith(".xlsx")).sorted().toList();
        }
    }
}
