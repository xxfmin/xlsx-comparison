package parity;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

import java.nio.file.*;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

class StatementParityTest {

    private static final Path EXPECTED =
        Path.of(System.getProperty("parity.expected", "src/test/resources/parity/expected"));
    private static final Path ACTUAL = Path.of(System.getProperty("parity.actual", "build/parity/actual"));

    static Stream<Arguments> corpus() throws Exception {
        // Skipped unless a corpus is configured, so a plain `mvn test` stays green; once it is
        // configured, an empty corpus fails rather than passing vacuously.
        assumeTrue(System.getProperty("parity.expected") != null,
            "set -Dparity.expected=<legacy dir> -Dparity.actual=<generated dir> to run the parity corpus");
        List<Path> files = xlsx(EXPECTED);
        assertFalse(files.isEmpty(), "Empty corpus proves nothing: " + EXPECTED);
        return files.stream().map(f -> Arguments.of(f.getFileName().toString(), f, ACTUAL.resolve(f.getFileName())));
    }

    /** A statement the new generator produces that legacy never did is also a parity failure. */
    @Test
    void generatorProducesNoUnexpectedStatements() throws Exception {
        assumeTrue(System.getProperty("parity.expected") != null, "no parity corpus configured");
        List<String> expected = xlsx(EXPECTED).stream().map(p -> p.getFileName().toString()).toList();
        List<String> extra = xlsx(ACTUAL).stream().map(p -> p.getFileName().toString())
            .filter(n -> !expected.contains(n)).toList();
        assertTrue(extra.isEmpty(), "Generated without a legacy counterpart: " + extra);
    }

    private static List<Path> xlsx(Path dir) throws Exception {
        if (!Files.isDirectory(dir)) return List.of();
        try (Stream<Path> s = Files.list(dir)) {
            return s.filter(f -> f.toString().toLowerCase().endsWith(".xlsx")).sorted().toList();
        }
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("corpus")
    void generatedStatementMatchesLegacy(String name, Path expected, Path actual) throws Exception {
        assertTrue(Files.exists(actual), "Not generated: " + actual);
        var result = WorkbookComparator.compare(expected, actual);
        assertTrue(result.identical(), () -> result.report(50));
    }
}
