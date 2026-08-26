using System.Text.Json;
using AIOrchestrator;
using AIOrchestrator.API;
using System.Diagnostics;

namespace SpreadsheetToolTests;

/// <summary>
/// SpreadsheetTool end-to-end test: the agent, given a natural-language task with rich
/// context (an Italian coffee-chain's H1 2026 data), must autonomously build a
/// professional multi-sheet XLSX workbook (data tables, formulas, charts, formatting)
/// using ONLY the SpreadsheetTool methods.
///
/// The outcome is NOT trusted from the agent's done message alone — the harness re-opens
/// the produced file with the tool itself, inspects sheets/charts/tables/formulas, and
/// then launches LibreOffice (installed on this machine) for a visual check.
/// </summary>
static class Program
{
    static int Main(string[] args)
    {
        // Method surface diagnostic: verifies which methods the LLM actually sees in the
        // generated tool definitions (GetToolDefinitions output) — no LLM involved.
        if (args.Contains("--methods"))
            return RunMethodsDiag();

        // Workbook inspection: --describe <path> opens an existing xlsx with the tool and
        // prints the agent's view (DescribeWorksheet formulaPatterns per sheet).
        var describeIdx = Array.FindIndex(args, a => string.Equals(a, "--describe", StringComparison.OrdinalIgnoreCase));
        if (describeIdx >= 0 && describeIdx + 1 < args.Length)
            return RunDescribe(args[describeIdx + 1]);

        // Deterministic chart/width check (no LLM): builds a workbook, adds charts, and
        // verifies the saved XML has populated chart caches and auto-fitted column widths.
        if (args.Contains("--charttest"))
            return RunChartTest();

        // Deterministic auto-format check (no LLM): default-width columns are fitted, user-set
        // widths preserved, and title cells (vertical runs above numbers/formulas) get a pastel
        // fill + bold, one color per table, without overriding explicit styles.
        if (args.Contains("--autofmt"))
            return RunAutoFormatTest();

        // Provider selection: --provider <name> (default DeepSeekBridge). Key-based providers
        // (DeepSeek/Zai/Gemini) read credentials from the per-app setup.json or the debug
        // preset — see Setup.Load/LoadDebugPreset docs.
        var providerName = "DeepSeekBridge";
        for (int i = 0; i < args.Length - 1; i++)
            if (string.Equals(args[i], "--provider", StringComparison.OrdinalIgnoreCase))
                providerName = args[i + 1];

        Console.WriteLine("╔══════════════════════════════════════════════════════╗");
        Console.WriteLine($"║  SpreadsheetTool test — Aurora Coffee H1 2026        ║");
        Console.WriteLine($"║  provider: {providerName,-45} ║");
        Console.WriteLine("╚══════════════════════════════════════════════════════╝");

        Log.IsEnabled = true;
        Log.LogStep($"=== SpreadsheetTool.Tests run (Aurora Coffee multi-sheet workbook, provider={providerName}) ===");

        var isBridge = string.Equals(providerName, "DeepSeekBridge", StringComparison.OrdinalIgnoreCase);
        if (isBridge)
        {
            if (!BridgeHealthy())
            {
                Console.WriteLine("\n✗ DeepseekBridge (127.0.0.1:8787) is not responding with real content.");
                Console.WriteLine("  Restart the bridge, then re-run this harness.");
                return 2;
            }
        }
        Setup.ProviderConfig = ProviderConfigs.Get(providerName);

        if (!isBridge)
        {
            Setup.Load();              // %LocalAppData%\SpreadsheetTool.Tests\setup.json
            Setup.LoadDebugPreset();   // debug_setup.json next to the executable
            if (string.IsNullOrEmpty(Setup.ApiKey))
            {
                Console.WriteLine($"\n✗ Provider '{providerName}' has no API key configured.");
                Console.WriteLine($"  Create '{Path.Combine(AppContext.BaseDirectory, "debug_setup.json")}' (dev-local, gitignored)");
                Console.WriteLine("  with e.g. {\"DeepSeekApiKey\": \"...\"} (or ZaiApiKey / GeminiApiKey), then re-run.");
                return 2;
            }
        }

        // Deterministic sandbox: a dedicated workspace folder owned by this harness.
        // AppContext.BaseDirectory = bin/Debug/net10.0/ → three levels up = project root.
        var workspace = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "workspace"));
        if (Directory.Exists(workspace)) Directory.Delete(workspace, recursive: true);
        Directory.CreateDirectory(workspace);
        Console.WriteLine($"\nWorkspace (sandbox root): {workspace}");

        // Skip the shared document index entirely (no document search needed here).
        Setup.SkipIndexingOnStartup = true;
        Setup.DocumentsPath = workspace;

        // Register the tool BEFORE the agent loop: McpToolRegistry resolves tool names by
        // scanning the loaded assemblies, and a referenced assembly loads lazily — without
        // this the catalog is empty and the agent invents method names.
        McpToolRegistry.Register(typeof(SpreadsheetTool));

        var orch = new AgentHarness(providerName);
        try
        {
            using var done = new ManualResetEventSlim();
            orch.AgentProgress += (_, e) =>
            {
                var line = e.State switch
                {
                    AgentHarness.AgentState.Iteration => $"[{e.TotalElapsedMs / 1000,4}s] iter {e.Iteration,2} → {e.MethodName}",
                    AgentHarness.AgentState.Completed => $"[{e.TotalElapsedMs / 1000,4}s] COMPLETED after {e.Iteration} iterations",
                    AgentHarness.AgentState.Failed => $"[{e.TotalElapsedMs / 1000,4}s] FAILED: {e.Error}",
                    _ => $"[{e.TotalElapsedMs / 1000,4}s] {e.State}"
                };
                Console.WriteLine(line);
                if (e.State is AgentHarness.AgentState.Completed or AgentHarness.AgentState.Failed)
                    done.Set();
            };

            Console.WriteLine("Starting agent scenario — building the Aurora Coffee workbook…");
            var task = Task.Run(() => orch.ExecuteAction(
                AuroraPrompt,
                new[] { "SpreadsheetTool" },
                maxIterations: 80));
            done.Wait(TimeSpan.FromSeconds(30)); // guard only — GetResult waits for completion
            var result = task.GetAwaiter().GetResult();

            Console.WriteLine($"\nAgent final message:\n{result.Message}\n");
            Console.WriteLine(result.Error != null ? $"Agent reported error: {result.Error}" : "");
            Console.WriteLine($"Outcome: {(result.Success ? "✓ SUCCESS" : $"✗ {result.Code}")} ({result.Iterations} iterations, {result.TotalElapsedMs / 1000}s)");

            // ── Structural verification: re-open the produced file with the tool itself ──
            var produced = Directory.GetFiles(workspace, "*.xlsx", SearchOption.AllDirectories)
                .OrderByDescending(f => File.GetLastWriteTimeUtc(f))
                .FirstOrDefault();
            if (produced == null)
            {
                Console.WriteLine("\n✗ No .xlsx file was produced in the workspace.");
                return 1;
            }

            Console.WriteLine($"\nProduced file: {produced} ({new FileInfo(produced).Length} bytes)");
            var issues = InspectWorkbook(produced);
            foreach (var i in issues)
                Console.WriteLine($"  ✗ {i}");
            if (issues.Count == 0)
                Console.WriteLine("  ✓ all structural checks passed");

            // ── Visual check: open the workbook in LibreOffice ──
            // KNOWN LIMITATION: LibreOffice's OOXML chart import renders the chart frame and
            // axes but NOT the series for charts written by Aspose.Cells.FOSS — the same file
            // renders correctly in Office Online / other viewers. The structural XML checks
            // above are the authoritative gate; LibreOffice is only a convenience peek.
            var soffice = FindSoffice();
            if (soffice == null)
            {
                Console.WriteLine("\nLibreOffice not found — open the file manually for the visual check.");
            }
            else
            {
                Console.WriteLine($"\nOpening with LibreOffice: {soffice} (note: charts may appear blank here — verify in Office Online if needed)");
                Process.Start(new ProcessStartInfo(soffice, $"\"{produced}\"") { UseShellExecute = true });
            }

            Console.WriteLine($"\nFull agent log: {Log.CurrentLogFile}");
            return result.Success ? 0 : 1;
        }
        finally
        {
            orch.Dispose();
        }
    }

    /// <summary>Re-opens the workbook with SpreadsheetTool and reports its structure:
    /// sheets, charts, tables, formula patterns, defined names. Returns the list of
    /// failed checks (empty = all passed).</summary>
    static List<string> InspectWorkbook(string filePath)
    {
        var issues = new List<string>();
        using var ss = new SpreadsheetTool();

        Console.WriteLine("Workbook structure:");
        if (ss.Open(filePath) != "true")
        {
            issues.Add("SpreadsheetTool.Open failed on the produced file");
            return issues;
        }

        var sheets = ss.GetSheetNames();
        Console.WriteLine($"  sheets ({sheets.Length}): {string.Join(", ", sheets)}");
        if (sheets.Length < 4)
            issues.Add($"expected a multi-sheet workbook (4 sheets: Vendite/Riepilogo/Personale/Grafici), found {sheets.Length}");

        string? describe = null;
        try { describe = ss.DescribeWorksheet(); }
        catch (Exception ex) { issues.Add($"DescribeWorksheet threw: {ex.Message}"); }

        int totalCharts = 0, totalTables = 0, sheetsWithFormulas = 0;
        if (describe != null)
        {
            using var doc = JsonDocument.Parse(describe);
            var root = doc.RootElement;
            if (root.TryGetProperty("definedNames", out var dns) && dns.GetArrayLength() > 0)
                Console.WriteLine($"  defined names ({dns.GetArrayLength()}): {string.Join(", ", dns.EnumerateArray().Select(d => d.GetProperty("name").GetString()))}");
            if (root.TryGetProperty("sheets", out var sarr))
            {
                foreach (var s in sarr.EnumerateArray())
                {
                    var name = s.GetProperty("name").GetString();
                    var charts = s.TryGetProperty("charts", out var c) ? c.GetArrayLength() : 0;
                    var tables = s.TryGetProperty("tables", out var t) ? t.GetArrayLength() : 0;
                    var hasFormulas = s.TryGetProperty("formulaPatterns", out var fp) && fp.GetArrayLength() > 0;
                    var used = s.TryGetProperty("usedRange", out var u) ? u.GetString() : "?";
                    totalCharts += charts;
                    totalTables += tables;
                    if (hasFormulas) sheetsWithFormulas++;
                    Console.WriteLine($"    {name}: usedRange={used}, charts={charts}, tables={tables}, formulas={(hasFormulas ? "yes" : "no")}");
                }
            }
        }
        if (totalCharts < 2)
            issues.Add($"expected at least 2 charts, found {totalCharts}");
        if (totalTables < 1)
            issues.Add($"expected at least 1 Excel table, found {totalTables}");
        if (sheetsWithFormulas == 0)
            issues.Add("no formula patterns found on any sheet");

        var dns2 = ss.GetDefinedNames();
        Console.WriteLine(dns2 != null && dns2.Length > 1
            ? $"  defined names via GetDefinedNames ({dns2.Length - 1})"
            : "  defined names: none");
        return issues;
    }

    /// <summary>Opens an existing workbook with the tool and prints the agent's view:
    /// sheets, used ranges and formula patterns per sheet (no LLM involved).</summary>
    static int RunDescribe(string filePath)
    {
        Console.WriteLine("╔══════════════════════════════════════════════════╗");
        Console.WriteLine("║  SpreadsheetTool workbook inspection (--describe) ║");
        Console.WriteLine("╚══════════════════════════════════════════════════╝");
        Log.IsEnabled = true;
        // Open() resolves paths against Setup.DocumentsPath: the sandbox root must be the
        // file's own folder, otherwise the absolute path is rejected as escaping the sandbox.
        Setup.SkipIndexingOnStartup = true;
        Setup.DocumentsPath = Path.GetDirectoryName(Path.GetFullPath(filePath)) ?? AppContext.BaseDirectory;
        using var ss = new SpreadsheetTool();
        if (ss.Open(filePath) != "true")
        {
            Console.WriteLine($"✗ cannot open '{filePath}'");
            return 1;
        }
        Console.WriteLine($"Sheets: {string.Join(", ", ss.GetSheetNames())}");
        try
        {
            using var doc = JsonDocument.Parse(ss.DescribeWorksheet());
            foreach (var s in doc.RootElement.GetProperty("sheets").EnumerateArray())
            {
                var name = s.GetProperty("name").GetString();
                var used = s.TryGetProperty("usedRange", out var u) ? u.GetString() : "?";
                Console.WriteLine($"\n{name}: usedRange={used}");
                if (s.TryGetProperty("formulaPatterns", out var fps) && fps.GetArrayLength() > 0)
                    foreach (var fp in fps.EnumerateArray())
                        Console.WriteLine($"    formula: {fp.GetProperty("formula").GetString()} [{fp.GetProperty("range").GetString()}]");
                else
                    Console.WriteLine("    formulas: none detected");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"✗ DescribeWorksheet failed: {ex.Message}");
            return 1;
        }
        return 0;
    }

    /// <summary>Deterministic verification (no LLM) of the two tool-side assists: charts must
    /// be saved with POPULATED caches (numCache points), and fresh columns must get auto
    /// width on first write. Inspects the produced xlsx XML directly.</summary>
    static int RunChartTest()
    {
        Console.WriteLine("╔══════════════════════════════════════════════════╗");
        Console.WriteLine("║  SpreadsheetTool chart/width check (--charttest) ║");
        Console.WriteLine("╚══════════════════════════════════════════════════╝");
        Log.IsEnabled = true;
        var dir = Path.Combine(Path.GetTempPath(), "SpreadsheetChartTest_" + Guid.NewGuid().ToString("N")[..6]);
        Directory.CreateDirectory(dir);
        Setup.SkipIndexingOnStartup = true;
        Setup.DocumentsPath = dir;

        var guardFail = false;
        using (var ss = new SpreadsheetTool())
        {
            ss.Create("/charttest.xlsx");
            ss.SetRange("Sheet1", "F1", new[] {
                new[]{"Mese","Torino","Milano","Bologna"},
                new[]{"Gen","100","200","300"},
                new[]{"Feb","150","250","350"},
                new[]{"Mar","180","280","380"},
            });
            // 2D block (4x4) → 3 series (Torino/Milano/Bologna) with categories Gen/Feb/Mar.
            var i1 = ss.AddChart("Sheet1", "Column", "Sheet1!$F$1:$I$4", 2, 6, 15, 20);
            // 2 columns → 1 series (Torino) with categories Gen/Feb/Mar (labels + values).
            var i2 = ss.AddChart("Sheet1", "Pie", "F1:G4", 17, 6, 30, 20);
            Console.WriteLine($"AddChart (anchored, normalized range): {i1}, {i2}");
            foreach (var info in ss.GetChartsInfo("Sheet1").Skip(1))
                Console.WriteLine($"  chart {info[0]}: {info[2]} at {info[3]}");

            // Guards: invalid parameters must return "error: …" (never throw, never a bare "false").
            var g1 = ss.SetRange("Sheet1", "A0", new[] { new[] { "x" } });
            var g2 = ss.AddChart("Sheet1", "Piee", "F1:G4", 0, 0, 5, 5);
            var g3 = ss.SetCellValue("Nope", "A1", "x");
            var g4 = ss.SetRange("Sheet1", "A1", new[] { new[] { "1", "2" }, new[] { "3" } });
            // Ragged block with a null row: must not throw (null row skipped, wider row auto-fitted).
            var g5 = ss.SetRange("Sheet1", "B6", new[] { new[] { "1", "2" }, null, new[] { "3", "4", "5" } });
            Console.WriteLine($"guards: SetRange(A0)='{g1}', AddChart(Piee)='{g2}', SetCellValue(Nope)='{g3}', SetRange(huge)='{g4}', SetRange(ragged+null)='{g5}'");
            guardFail = !g1.StartsWith("Error:") || !g2.StartsWith("Error:") || !g3.StartsWith("Error:") || g4.StartsWith("Error:") || g5.StartsWith("Error:");
            if (guardFail) Console.WriteLine("  ✗ a guard returned something other than 'error: …' on invalid input");
        }

        var file = Path.Combine(dir, "charttest.xlsx");
        var ok = guardFail == false;
        using (var zip = System.IO.Compression.ZipFile.OpenRead(file))
        {
            var chartEntries = zip.Entries.Where(e => e.FullName.StartsWith("xl/charts/") && e.FullName.EndsWith(".xml")).ToList();
            if (chartEntries.Count != 2)
            { ok = false; Console.WriteLine($"  ✗ expected 2 chart parts, found {chartEntries.Count}"); }

            foreach (var entry in chartEntries)
            {
                using var sr = new StreamReader(entry.Open());
                var xml = sr.ReadToEnd();

                // Well-formedness is the SAME bar Excel/LibreOffice apply on open: a malformed
                // chart part is a blank chart even when the caches look populated.
                var doc = new System.Xml.XmlDocument();
                try { doc.LoadXml(xml); }
                catch (System.Xml.XmlException ex)
                { ok = false; Console.WriteLine($"  ✗ {entry.FullName} is NOT well-formed: {ex.Message}"); }

                var sers = System.Text.RegularExpressions.Regex.Matches(xml, "<c:ser>").Count;
                var cats = System.Text.RegularExpressions.Regex.Matches(xml, "<c:cat>").Count;
                var cached = System.Text.RegularExpressions.Regex.Matches(xml, "<c:v>").Count;
                var pt = System.Text.RegularExpressions.Regex.Match(xml, "<c:ptCount val=\"(\\d+)\"").Groups[1].Value;
                Console.WriteLine($"{entry.FullName}: series={sers}, categories={cats}, <c:v>={cached}, ptCount={pt}");
                if (cached == 0) { ok = false; Console.WriteLine("  ✗ chart cache EMPTY — would render blank in LibreOffice"); }
            }

            // chart1 (Column, 4x4 matrix) must be restructured: 3 series + 3 categories.
            var chart1 = zip.GetEntry("xl/charts/chart1.xml");
            if (chart1 != null)
            {
                using var sr = new StreamReader(chart1.Open());
                var xml = sr.ReadToEnd();
                var sers = System.Text.RegularExpressions.Regex.Matches(xml, "<c:ser>").Count;
                var cats = System.Text.RegularExpressions.Regex.Matches(xml, "<c:cat>").Count;
                if (sers != 3) { ok = false; Console.WriteLine($"  ✗ chart1: expected 3 series (matrix 4x4), found {sers}"); }
                if (cats != 3) { ok = false; Console.WriteLine($"  ✗ chart1: expected 3 categories, found {cats}"); }
            }
            // chart2 (Pie, 2 columns) → 1 series with 1 category set.
            var chart2 = zip.GetEntry("xl/charts/chart2.xml");
            if (chart2 != null)
            {
                using var sr = new StreamReader(chart2.Open());
                var xml = sr.ReadToEnd();
                var sers = System.Text.RegularExpressions.Regex.Matches(xml, "<c:ser>").Count;
                var cats = System.Text.RegularExpressions.Regex.Matches(xml, "<c:cat>").Count;
                if (sers != 1) { ok = false; Console.WriteLine($"  ✗ chart2: expected 1 series (2-column range), found {sers}"); }
                if (cats != 1) { ok = false; Console.WriteLine($"  ✗ chart2: expected 1 category set, found {cats}"); }
            }

            var sheetXml = zip.GetEntry("xl/worksheets/sheet1.xml");
            if (sheetXml != null)
            {
                using var sr = new StreamReader(sheetXml.Open());
                var xml = sr.ReadToEnd();
                var bestFit = System.Text.RegularExpressions.Regex.Matches(xml, "<col [^>]*bestFit=\"1\"");
                var widthCols = System.Text.RegularExpressions.Regex.Matches(xml, "<col [^>]*width=\"([0-9.]+)\"");
                Console.WriteLine($"sheet1 bestFit columns: {bestFit.Count}, explicit-width columns: {widthCols.Count}");
                if (bestFit.Count == 0) { ok = false; Console.WriteLine("  ✗ no bestFit columns written"); }
            }
        }
        Console.WriteLine(ok ? "\n✓ chart caches populated + series restructured + XML well-formed" : "\n✗ deterministic assists failed");
        return ok ? 0 : 1;
    }

    /// <summary>Deterministic auto-format check (--autofmt, no LLM): builds a workbook with two
    /// tables (general title + column headers + numeric data + formulas + a number format), a
    /// user-set column width and an explicitly styled header row; then verifies the saved XML:
    /// default-width columns got fitted widths (format-aware), the user width is preserved, title
    /// cells (vertical runs included) got a pastel fill + bold, the two tables got different
    /// colors, and the FormatHeaderRow blue was not overridden.</summary>
    static int RunAutoFormatTest()
    {
        Console.WriteLine("╔══════════════════════════════════════════════════╗");
        Console.WriteLine("║  SpreadsheetTool deterministic auto-format check ║");
        Console.WriteLine("╚══════════════════════════════════════════════════╝");
        Log.IsEnabled = true;
        var dir = Path.Combine(Path.GetTempPath(), "SpreadsheetAutoFmt_" + Guid.NewGuid().ToString("N")[..6]);
        Directory.CreateDirectory(dir);
        Setup.SkipIndexingOnStartup = true;
        Setup.DocumentsPath = dir;

        using (var ss = new SpreadsheetTool())
        {
            ss.Create("/autofmt.xlsx");
            // Table 1: general title (A1), headers (row 2), numeric data (incl. Year in col A) +
            // Profit formulas; currency format on B:D; FormatHeaderRow explicitly styles row 0.
            ss.SetCellValue("Sheet1", "A1", "Coffee Shop Report");
            ss.SetRange("Sheet1", "A2", new[]
            {
                new[] { "Year", "Revenue", "Costs", "Profit" },
                new[] { "2023", "12500", "8400", "=B3-C3" },
                new[] { "2024", "13200", "9100", "=B4-C4" },
                new[] { "2025", "11800", "7900", "=B5-C5" },
            });
            // Currency format passed WITH surrounding quotes (as models often emit it): the tool
            // must normalize it — a quoted string is a LITERAL in Excel format syntax.
            ss.ApplyStyle("Sheet1", "B3:D5", numberFormat: "\"$#,##0.00\"");
            ss.FormatHeaderRow("Sheet1");          // blue on A1 — must NOT be overridden
            // Table 2: general title (A7), headers (row 8), numeric data.
            ss.SetCellValue("Sheet1", "A7", "Quarterly Targets");
            ss.SetRange("Sheet1", "A8", new[]
            {
                new[] { "Q1", "Q2" },
                new[] { "100", "120" },
                new[] { "110", "130" },
            });
            // User-set width on column E: must be preserved by the auto-format pass.
            ss.SetColumnWidth("Sheet1", 4, 20);
            ss.SetCellValue("Sheet1", "E2", "wide user column");
            Console.WriteLine("  scenario: Create + SetRange + ApplyStyle + FormatHeaderRow + SetColumnWidth → Save");
            Console.WriteLine($"  Save: {ss.Save()}");
        }

        var file = Path.Combine(dir, "autofmt.xlsx");
        var ok = true;
        using (var zip = System.IO.Compression.ZipFile.OpenRead(file))
        {
            var sheet = ReadEntry(zip, "xl/worksheets/sheet1.xml");

            // 1) bestFit on the touched default-width columns (NO width value); the user-set
            //    width on column E is preserved and must NOT get bestFit.
            var colsSection = System.Text.RegularExpressions.Regex.Match(sheet, "<cols>.*?</cols>").Value;
            Console.WriteLine($"  cols: {colsSection}");
            var bestFitCols = new HashSet<int>();
            foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(colsSection,
                "<col min=\"(\\d+)\"[^>]*bestFit=\"1\"[^>]*/>"))
                bestFitCols.Add(int.Parse(m.Groups[1].Value));
            var userWidths = new Dictionary<int, double>();
            foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(colsSection,
                "<col min=\"(\\d+)\"[^>]*width=\"([0-9.]+)\""))
                userWidths[int.Parse(m.Groups[1].Value)] = double.Parse(m.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);

            // A,B,C,D written by the agent with default width → bestFit, no width value.
            foreach (var c in new[] { 1, 2, 3, 4 })
                if (!bestFitCols.Contains(c)) { ok = false; Console.WriteLine($"  ✗ col {ColLetter(c)} expected bestFit"); }
            if (userWidths.ContainsKey(1) || userWidths.ContainsKey(2) || userWidths.ContainsKey(3) || userWidths.ContainsKey(4))
            { ok = false; Console.WriteLine("  ✗ a bestFit column must not carry a width value"); }
            // E has a user-set width (20) → preserved, no bestFit.
            if (!userWidths.TryGetValue(5, out var wE) || Math.Abs(wE - 20) > 0.01) { ok = false; Console.WriteLine($"  ✗ col E expected width 20 (user width preserved), got {wE}"); }
            if (bestFitCols.Contains(5)) { ok = false; Console.WriteLine("  ✗ col E must NOT get bestFit (user width set)"); }

            // 3) Number format normalization: the quoted "$#,##0.00" must be saved WITHOUT quotes.
            var styles = ReadEntry(zip, "xl/styles.xml");
            var formatCodes = new List<string>();
            foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(styles,
                "<numFmt[^>]*formatCode=\"([^\"]*)\""))
                formatCodes.Add(System.Net.WebUtility.HtmlDecode(m.Groups[1].Value));
            Console.WriteLine($"  numFmt codes: {string.Join(" | ", formatCodes)}");
            if (formatCodes.Any(fc => fc.Contains('"'))) { ok = false; Console.WriteLine("  ✗ a number format still contains literal quotes (would render as text)"); }
            if (!formatCodes.Contains("$#,##0.00")) { ok = false; Console.WriteLine("  ✗ quoted \"$#,##0.00\" was not normalized to $#,##0.00"); }

            // 2) Title styling: fill color per cell, mapped through styles.xml.
            var fillByStyle = ParseFillByStyle(styles);
            var cellFill = new Dictionary<string, string>();
            foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(sheet,
                "<c r=\"([A-Z]+\\d+)\"(?:[^>]*?s=\"(\\d+)\")?"))
            {
                var refName = m.Groups[1].Value;
                var styleIdx = m.Groups[2].Success ? int.Parse(m.Groups[2].Value) : 0;
                cellFill[refName] = fillByStyle.TryGetValue(styleIdx, out var f) ? f : "(none)";
            }
            string F(string cell) => cellFill.TryGetValue(cell, out var f) ? f : "(none)";

            Console.WriteLine($"  fills: A1={F("A1")} A2={F("A2")} A7={F("A7")} A8={F("A8")} A9={F("A9")} B8={F("B8")}");
            if (!F("A1").Equals("FF2278D4", StringComparison.OrdinalIgnoreCase)) { ok = false; Console.WriteLine("  ✗ A1 (FormatHeaderRow blue) was overridden"); }
            if (F("A2") == "(none)" || F("A2") == "FF2278D4") { ok = false; Console.WriteLine("  ✗ A2 title not pastel-styled"); }
            if (F("A7") != F("A8")) { ok = false; Console.WriteLine("  ✗ A7 general title must share the table-2 color with A8"); }
            if (F("A2") == F("A7")) { ok = false; Console.WriteLine("  ✗ table 1 and table 2 must get different colors"); }
            if (F("A9") != "(none)") { ok = false; Console.WriteLine($"  ✗ A9 (data) must not be styled, got {F("A9")}"); }
            if (F("B8") != F("A8")) { ok = false; Console.WriteLine("  ✗ B8 header must share the table-2 color"); }
        }
        Console.WriteLine(ok ? "\n✓ auto-format deterministic checks passed" : "\n✗ auto-format deterministic checks failed");
        return ok ? 0 : 1;
    }

    static string ReadEntry(System.IO.Compression.ZipArchive zip, string name)
    {
        var entry = zip.GetEntry(name);
        if (entry == null) return "";
        using var sr = new StreamReader(entry.Open());
        return sr.ReadToEnd();
    }

    /// <summary>Maps a cellXfs style index to its fill color rgb (or "(none)"). Parses the
    /// fills list and the cellXfs fillId references from styles.xml.</summary>
    static Dictionary<int, string> ParseFillByStyle(string stylesXml)
    {
        var fills = new List<string>();
        foreach (System.Text.RegularExpressions.Match fm in System.Text.RegularExpressions.Regex.Matches(stylesXml, "<fill>(.*?)</fill>"))
        {
            var rgb = System.Text.RegularExpressions.Regex.Match(fm.Groups[1].Value, "fgColor rgb=\"([0-9A-Fa-f]{8})\"");
            fills.Add(rgb.Success ? rgb.Groups[1].Value.ToUpperInvariant() : "(none)");
        }
        var result = new Dictionary<int, string>();
        var xfs = System.Text.RegularExpressions.Regex.Match(stylesXml, "<cellXfs[^>]*>(.*?)</cellXfs>", System.Text.RegularExpressions.RegexOptions.Singleline);
        if (!xfs.Success) return result;
        int idx = 0;
        foreach (System.Text.RegularExpressions.Match x in System.Text.RegularExpressions.Regex.Matches(xfs.Groups[1].Value, "<xf[^>]*>"))
        {
            var fillId = System.Text.RegularExpressions.Regex.Match(x.Value, "fillId=\"(\\d+)\"").Groups[1].Value;
            var f = int.TryParse(fillId, out var fi) && fi < fills.Count ? fills[fi] : "(none)";
            result[idx++] = f;
        }
        return result;
    }

    static string ColLetter(int col) => col >= 1 && col <= 26 ? ((char)('A' + col - 1)).ToString() : "?";

    /// <summary>Verifies the method surface the LLM is given: GetToolDefinitions must
    /// contain every method the scenario relies on. No LLM involved — this proves whether
    /// an agent's "method X doesn't exist" claim is a tool gap or a model hallucination.</summary>
    static int RunMethodsDiag()
    {
        Console.WriteLine("╔══════════════════════════════════════════════════╗");
        Console.WriteLine("║  SpreadsheetTool method surface diagnostic       ║");
        Console.WriteLine("╚══════════════════════════════════════════════════╝");
        Log.IsEnabled = true;

        var defs = UISupportGeneric.Analyzer.GetToolDefinitions(typeof(SpreadsheetTool));
        var required = new[]
        {
            "create", "save", "save_as", "get_sheet_names", "rename_worksheet", "add_worksheet",
            "describe_worksheet", "set_range", "set_cell_formula", "add_table", "set_auto_filter",
            "add_conditional_format", "add_chart", "add_defined_name", "apply_style"
        };
        var missing = required.Where(m => !defs.Contains($"### {m}", StringComparison.Ordinal)).ToList();
        Console.WriteLine($"Tool definitions: {defs.Length} chars");
        Console.WriteLine(missing.Count == 0
            ? "✓ all expected methods present in the LLM-facing definitions"
            : $"✗ MISSING from definitions: {string.Join(", ", missing)}");

        Console.WriteLine("\nMethods exposed to the LLM:");
        foreach (var line in defs.Split('\n'))
            if (line.TrimStart().StartsWith("### "))
                Console.WriteLine("  " + line.Trim());

        return missing.Count == 0 ? 0 : 1;
    }

    /// <summary>Locates the LibreOffice executable on this Windows machine.</summary>
    static string? FindSoffice()
    {
        var candidates = new[]
        {
            @"C:\Program Files\LibreOffice\program\soffice.exe",
            @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
        };
        return candidates.FirstOrDefault(File.Exists);
    }

    /// <summary>Liveness check for the DeepseekBridge (127.0.0.1:8787): a minimal
    /// request must return at least one chunk with non-empty content. An HTTP 200
    /// with only empty deltas means the web quota is throttled (known failure mode).</summary>
    static bool BridgeHealthy()
    {
        try
        {
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
            var body = JsonSerializer.Serialize(new
            {
                model = "deepseek-web/deepseek-chat",
                messages = new[] { new { role = "user", content = "Say banana" } },
                max_tokens = 10
            });
            var resp = client.PostAsync("http://127.0.0.1:8787/v1/chat/completions",
                new StringContent(body, System.Text.Encoding.UTF8, "application/json")).GetAwaiter().GetResult();
            var text = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            foreach (var line in text.Split('\n'))
            {
                if (line.Contains("\"delta\":{\"content\":\"") && !line.Contains("\"delta\":{\"content\":\"\"}"))
                    return true;
            }
            Console.WriteLine("Bridge response had no non-empty content chunk (throttled?).");
            return false;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Bridge health check failed: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Scenario prompt: the agent must build a professional multi-sheet XLSX workbook for
    /// "Aurora Coffee" (chain of 3 coffee shops) with H1 2026 data — detail table,
    /// month×shop matrix, Excel table, autofilter, conditional formatting, cross-sheet KPI
    /// formulas, defined names, staff costs, and 2 charts. The agent discovers the tool
    /// surface via the provided method definitions and is expected to report any blocker
    /// honestly (e.g. no method to add worksheets).
    /// </summary>
    const string AuroraPrompt = """
        Crea un workbook Excel professionale e sofisticato per la catena "Aurora Coffee" con i dati del primo semestre 2026.

        CONTESTO AZIENDALE:
        Aurora Coffee S.r.l. gestisce 3 caffetterie: "Torino Centro", "Milano Navigli", "Bologna D'Azeglio".
        Vendono 3 categorie: Caffetteria, Pasticceria, Merchandising.
        Questo workbook servirà alla direzione per analizzare l'andamento del semestre e presentarlo al consiglio: deve essere completo, formattato bene e con grafici.

        DATI (primo semestre 2026, valori in euro):

        VENDITE DETTAGLIO (Mese; Locale; Categoria; Importo):
        Gennaio;Torino Centro;Caffetteria;14500
        Gennaio;Torino Centro;Pasticceria;6200
        Gennaio;Torino Centro;Merchandising;2100
        Gennaio;Milano Navigli;Caffetteria;17800
        Gennaio;Milano Navigli;Pasticceria;8300
        Gennaio;Milano Navigli;Merchandising;3400
        Gennaio;Bologna D'Azeglio;Caffetteria;12100
        Gennaio;Bologna D'Azeglio;Pasticceria;5400
        Gennaio;Bologna D'Azeglio;Merchandising;1900
        Febbraio;Torino Centro;Caffetteria;15200
        Febbraio;Torino Centro;Pasticceria;6800
        Febbraio;Torino Centro;Merchandising;1800
        Febbraio;Milano Navigli;Caffetteria;18100
        Febbraio;Milano Navigli;Pasticceria;8700
        Febbraio;Milano Navigli;Merchandising;3600
        Febbraio;Bologna D'Azeglio;Caffetteria;12800
        Febbraio;Bologna D'Azeglio;Pasticceria;5900
        Febbraio;Bologna D'Azeglio;Merchandising;2100
        Marzo;Torino Centro;Caffetteria;14800
        Marzo;Torino Centro;Pasticceria;7100
        Marzo;Torino Centro;Merchandising;2400
        Marzo;Milano Navigli;Caffetteria;17500
        Marzo;Milano Navigli;Pasticceria;9100
        Marzo;Milano Navigli;Merchandising;3200
        Marzo;Bologna D'Azeglio;Caffetteria;12600
        Marzo;Bologna D'Azeglio;Pasticceria;6100
        Marzo;Bologna D'Azeglio;Merchandising;2300
        Aprile;Torino Centro;Caffetteria;16300
        Aprile;Torino Centro;Pasticceria;6900
        Aprile;Torino Centro;Merchandising;2600
        Aprile;Milano Navigli;Caffetteria;19200
        Aprile;Milano Navigli;Pasticceria;8800
        Aprile;Milano Navigli;Merchandising;4100
        Aprile;Bologna D'Azeglio;Caffetteria;13900
        Aprile;Bologna D'Azeglio;Pasticceria;6300
        Aprile;Bologna D'Azeglio;Merchandising;2400
        Maggio;Torino Centro;Caffetteria;17100
        Maggio;Torino Centro;Pasticceria;7400
        Maggio;Torino Centro;Merchandising;2900
        Maggio;Milano Navigli;Caffetteria;20400
        Maggio;Milano Navigli;Pasticceria;9400
        Maggio;Milano Navigli;Merchandising;4700
        Maggio;Bologna D'Azeglio;Caffetteria;14500
        Maggio;Bologna D'Azeglio;Pasticceria;6600
        Maggio;Bologna D'Azeglio;Merchandising;2600
        Giugno;Torino Centro;Caffetteria;16900
        Giugno;Torino Centro;Pasticceria;7800
        Giugno;Torino Centro;Merchandising;3100
        Giugno;Milano Navigli;Caffetteria;21100
        Giugno;Milano Navigli;Pasticceria;10200
        Giugno;Milano Navigli;Merchandising;5200
        Giugno;Bologna D'Azeglio;Caffetteria;14100
        Giugno;Bologna D'Azeglio;Pasticceria;7100
        Giugno;Bologna D'Azeglio;Merchandising;2800

        TOTALI MENSILI PER LOCALE (matrice già calcolata, usala per i grafici):
        Gennaio;Torino Centro;22800
        Gennaio;Milano Navigli;29500
        Gennaio;Bologna D'Azeglio;19400
        Febbraio;Torino Centro;23800
        Febbraio;Milano Navigli;30400
        Febbraio;Bologna D'Azeglio;20800
        Marzo;Torino Centro;24300
        Marzo;Milano Navigli;29800
        Marzo;Bologna D'Azeglio;21000
        Aprile;Torino Centro;25800
        Aprile;Milano Navigli;32100
        Aprile;Bologna D'Azeglio;22600
        Maggio;Torino Centro;27400
        Maggio;Milano Navigli;34500
        Maggio;Bologna D'Azeglio;23700
        Giugno;Torino Centro;27800
        Giugno;Milano Navigli;36500
        Giugno;Bologna D'Azeglio;24000

        TARGET MENSILE DI RICAVO PER LOCALE: Torino Centro 32000; Milano Navigli 40000; Bologna D'Azeglio 27000.

        PERSONALE (Locale; Ruolo; N. addetti; Costo mensile per addetto):
        Torino Centro;Barista;2;1800
        Torino Centro;Pasticciere;1;1900
        Torino Centro;Addetto sala part-time;1;1100
        Milano Navigli;Barista;3;1850
        Milano Navigli;Pasticciere;1;2000
        Milano Navigli;Addetto sala part-time;1;1200
        Bologna D'Azeglio;Barista;2;1750
        Bologna D'Azeglio;Pasticciere;1;1850

        TASK (esegui in ordine):
        1. Crea il workbook con Create("/aurora_coffee_2026.xlsx"). Poi GetSheetNames() e DescribeWorksheet() per scoprire la struttura iniziale.
        2. Costruisci il foglio "Vendite":
           - Tabella DETTAGLIO con intestazione (Mese; Locale; Categoria; Importo) e le 54 righe di dati del dettaglio.
           - Una MATRICE compatta Mese×Locale (colonne: Mese, Torino Centro, Milano Navigli, Bologna D'Azeglio; 6 righe mesi) con i TOTALI MENSILI forniti, più una riga "Target mensile" con i target, più una riga "Totale semestre" con formule =SUM della riga.
           - Applica alla matrice un formato condizionale (AddConditionalFormat, CellValue/GreaterThan) che evidenzi le celle del totale mese superiori al target del locale.
           - Sulla tabella dettaglio: tabella Excel (AddTable), autofilter (SetAutoFilter), header formattato (FormatHeaderRow), formato numero euro ("€ #,##0") sulle colonne importo, bordi sottili sul range dati, larghezze colonna ragionevoli.
        3. Costruisci il foglio "Riepilogo" con i KPI per la direzione, usando FORMULE che referenziano "Vendite" (e "Personale" dove serve): ricavi totali semestre; ricavi per locale; quota percentuale per locale; migliore mese per ricavi (es. INDEX/MATCH o MAX); margine lordo stimato = ricavi totali × MargineLordo (definisci il nome MargineLordo = 0.62 con AddDefinedName); costo personale mensile totale e incidenza percentuale sui ricavi. Formatta: titolo grande su celle unite, stile professionale (colori, bordi, valute, percentuali).
        4. Costruisci il foglio "Personale": tabella con i dati forniti, colonna "Costo mensile totale" con formula (N. addetti × Costo per addetto), totale per locale (formule SUMIF o riferimenti) e totale generale. Formatta header e valute.
        5. Costruisci il foglio "Grafici" con almeno 2 grafici tramite AddChart, referenziando la MATRICE Mese×Locale del foglio Vendite:
           - grafico a COLONNE: ricavi mensili per locale;
           - grafico a TORTA (pie): quota ricavi per locale sul semestre.
           Posizionali in aree diverse del foglio. Se AddChart richiede un range compatto, verifica con DescribeWorksheet dove sta la matrice e usa il range esatto (es. "Vendite!$A$2:$D$8").
        6. Salva con Save() e fai la verifica finale: GetSheetNames() e DescribeWorksheet() per confermare fogli, grafici, tabelle e formule. Poi rispondi con il metodo done ({"method": "done", "message": "..."}) riepilogando: percorso del file, fogli creati, grafici inseriti, formule principali, e le DIFFICOLTÀ incontrate durante l'uso del tool (metodi mancanti, errori, tentativi falliti e come li hai risolti).

        REGOLE DI USO DEL TOOL:
        - I percorsi file sono Unix-style, relativi alla workspace root (es. "/aurora_coffee_2026.xlsx"): non uscire MAI dalla sandbox.
        - Usa SOLO i metodi elencati nelle definizioni disponibili: non inventare nomi di metodi o parametri. I nomi foglio sono case-sensitive e vanno presi da GetSheetNames().
        - Le formule (SetCellFormula) vengono salvate ma NON ricalcolate dal tool: non tentare di verificarne il valore leggendolo, le ricalcolerà l'app di visualizzazione.
        - Se un metodo restituisce un errore, rileggi il messaggio e correggi i parametri; persisti 2-3 tentativi prima di cambiare strategia.
        - Se NON esiste alcun metodo per aggiungere fogli di lavoro al workbook, NON inventarlo: organizza tutto il lavoro in sezioni ben delimitate sul foglio disponibile (una per area, con titoli) e segnala ESPLICITAMENTE questo limite nel messaggio done.
        - Non usare il web né altri strumenti: hai solo SpreadsheetTool.
        """;
}
