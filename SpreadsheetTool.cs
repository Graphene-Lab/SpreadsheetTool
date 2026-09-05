using Aspose.Cells_FOSS;
using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;

namespace AIOrchestrator.API
{
    /// <summary>
    /// Spreadsheet (XLSX) operations for agent use: open/create, cells, ranges, styles, charts, tables.
    /// </summary>
    public class SpreadsheetTool : BaseAgentTool, IDisposable, IFileTool
    {
        private Workbook? _workbook;
        private string _filePath = string.Empty;

        /// <summary>True when the in-memory workbook differs from the file on disk. Lets Dispose
        /// skip the redundant second save when the agent already called Save explicitly.</summary>
        private bool _dirty;

        /// <summary>Columns the agent wrote to or styled this session, per worksheet. Only these
        /// get the bestFit flag at save: a column the USER set up (width or plain content) is
        /// never touched unless the agent works on it.</summary>
        private readonly Dictionary<string, HashSet<int>> _touchedCols = new();

        /// <summary>Columns that will receive the OOXML bestFit flag (auto width computed by the
        /// opening application) at save, per worksheet. Populated by the deterministic auto-format
        /// pass right before persisting.</summary>
        private readonly Dictionary<string, HashSet<int>> _bestFitCols = new();

        /// <summary>Guard for agent-supplied ranges: the tool must never let a huge range
        /// (e.g. "A1:XFD1048576") allocate unbounded memory or hang the session.</summary>
        private const int MaxCellArea = 1_000_000;

        /// <summary>
        /// Parameterless constructor for agent activation. Call <see cref="Open"/> or <see cref="Create"/>
        /// before using other methods.
        /// </summary>
        public SpreadsheetTool()
        {
        }

        /// <summary>
        /// Opens an existing XLSX workbook for editing.
        /// </summary>
        /// <param name="filePath">
        /// Path to an existing .xlsx file, Unix style relative to the workspace root (leading
        /// "/", e.g. "/folder/file.xlsx").
        /// </param>
        public SpreadsheetTool(string filePath)
        {
            Open(filePath);
        }

        /// <summary>
        /// Opens an existing XLSX workbook and replaces the current one.
        /// Call this when the agent already has an instance (created via parameterless constructor)
        /// and needs to load a specific file.
        /// </summary>
        /// <param name="filePath">Path to an existing .xlsx file (Unix style, e.g. "/folder/file.xlsx").</param>
        /// <returns>"true", or "Error: …" when the file cannot be opened.</returns>
        public string Open(string filePath)
        {
            try
            {
                _workbook?.Dispose();
                _filePath = SandboxPath.Resolve(filePath);
                _workbook = new Workbook(_filePath);
                _touchedCols.Clear();
                _bestFitCols.Clear();
                _dirty = false;
                Log.LogStep($"SpreadsheetTool.Open: opened '{_filePath}'");
                return "true";
            }
            catch (Exception ex)
            {
                Log.LogStep($"SpreadsheetTool.Open: failed '{filePath}': {ex.Message}");
                return $"Error: {ex.Message}";
            }
        }

        /// <summary>
        /// Creates a new XLSX workbook with one default worksheet ("Sheet1") on THIS instance.
        /// Must be an instance method (not a static factory): the agent loop keeps ONE shared
        /// instance in its agents dictionary, so a static Create returning a brand-new agent
        /// discarded the workbook and every later edit failed with a NullReferenceException.
        /// </summary>
        /// <param name="filePath">
        /// Path where the new .xlsx file will be saved, Unix style relative to the workspace
        /// root (e.g. "/folder/file.xlsx").
        /// </param>
        /// <returns>"true", or "Error: …" when the file cannot be created.</returns>
        public string Create(string filePath)
        {
            try
            {
                var resolved = SandboxPath.Resolve(filePath);
                _workbook?.Dispose();
                _workbook = new Workbook();
                _touchedCols.Clear();
                _bestFitCols.Clear();
                _workbook.Save(resolved);
                _filePath = resolved;
                _dirty = false;
                Log.LogStep($"SpreadsheetTool.Create: created '{resolved}'");
                return $"Workbook created at '{SandboxPath.ToAgent(resolved)}'.";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        /// <summary>
        /// Writes all pending changes to the current file path — an explicit checkpoint.
        /// (Changes are also persisted automatically when the tool is disposed, so the file
        /// on disk always reflects the final session state.)
        /// A save that would produce an unreadable file is rejected: the previous file is left
        /// untouched and no version is created.
        /// The new content becomes a new version in the workspace git repo (rollback via GitTool.restore).
        /// </summary>
        /// <returns>A message describing the result: the new version id, or an error when the
        /// save was rejected.</returns>
        public string Save()
        {
            if (_workbook == null) return "No changes to save — no workbook is open.";
            if (!_dirty) return "No changes to save — the workbook is unchanged since the last save.";

            ApplyDeterministicAutoFormat();
            var validationError = PersistValidated(_filePath);
            if (validationError != null)
            {
                Log.LogStep($"SpreadsheetTool.Save: REJECTED — {validationError}");
                return "Error: save rejected — the workbook could not be saved correctly; the previous file and the last git version are untouched.";
            }

            var versionId = GitSupport.Snapshot(_filePath, "SpreadsheetTool save");
            _dirty = false;
            Log.LogStep($"SpreadsheetTool.Save: saved to '{_filePath}', version='{versionId}'");
            var agentPath = SandboxPath.ToAgent(_filePath);
            return versionId != null
                ? $"Workbook saved to '{agentPath}'. New version: {versionId}. (Rollback via GitTool.restore.)"
                : $"Workbook saved to '{agentPath}'. (No changes detected.)";
        }

        /// <summary>
        /// Writes all pending changes to a new file path.
        /// Subsequent Save() calls will use the new path.
        /// A save that would produce an unreadable file is rejected: the target file is not created.
        /// The new content becomes a new version in the workspace git repo.
        /// </summary>
        /// <param name="newFilePath">Path for the new .xlsx file, Unix style relative to the workspace root (e.g. "/folder/file.xlsx").</param>
        /// <returns>A message describing the result, or an error when the save was rejected.</returns>
        public string SaveAs(string newFilePath)
        {
            if (_workbook == null) return "No changes to save — no workbook is open.";

            var resolved = SandboxPath.Resolve(newFilePath);
            ApplyDeterministicAutoFormat();
            var validationError = PersistValidated(resolved);
            if (validationError != null)
            {
                Log.LogStep($"SpreadsheetTool.SaveAs: REJECTED — {validationError}");
                return "Error: save rejected — the workbook could not be saved correctly; the target file was not created.";
            }

            _filePath = resolved;
            var versionId = GitSupport.Snapshot(_filePath, "SpreadsheetTool save as");
            _dirty = false;
            Log.LogStep($"SpreadsheetTool.SaveAs: saved to '{_filePath}', version='{versionId}'");
            var agentPath = SandboxPath.ToAgent(resolved);
            return versionId != null
                ? $"Workbook saved as '{agentPath}'. New version: {versionId}."
                : $"Workbook saved as '{agentPath}'.";
        }

        /// <summary>Reverts the OPEN workbook to a version from the workspace git repo (list them with
        /// GitTool.history). The current state is saved as a new version first (the rollback is
        /// reversible), then the file is overwritten and the workbook is reloaded. Use this when the
        /// workbook is open in this tool; GitTool.restore handles files that are not open.</summary>
        /// <param name="versionId">Version to restore, from GitTool.history().</param>
        /// <returns>Descriptive result message.</returns>
        public string Restore(string versionId)
        {
            if (_workbook == null) return "No workbook is open. Nothing to restore.";
            try
            {
                _workbook.Dispose();   // release the open handle so the file can be overwritten
                var message = GitSupport.Restore(versionId, _filePath);
                _workbook = new Workbook(_filePath);
                _dirty = false;        // reloaded state matches the file on disk
                return message;
            }
            catch (Exception ex)
            {
                // Never leave the tool with a null workbook: reload the file (git restores it
                // atomically per file) or fall back to a blank workbook so later calls keep
                // working; the agent still gets a descriptive error.
                try { if (_workbook == null) _workbook = new Workbook(_filePath); }
                catch { _workbook = new Workbook(); _filePath = string.Empty; }
                _dirty = true;
                return $"Error: Restore failed: {ex.Message}";
            }
        }

        /// <summary>
        /// Explicit interface implementation — NOT an agent tool (the orchestrator disposes
        /// agents automatically when the loop ends). Persists any unsaved changes first so the
        /// file on disk always reflects the final session state, then releases the workbook.
        /// A save that fails is not committed: the last good file stays on disk.
        /// </summary>
        void IDisposable.Dispose()
        {
            try
            {
                if (_workbook != null && !string.IsNullOrEmpty(_filePath) && _dirty)
                {
                    ApplyDeterministicAutoFormat();
                    var validationError = PersistValidated(_filePath);
                    if (validationError != null)
                    {
                        Log.LogStep($"SpreadsheetTool.Dispose: auto-save REJECTED — {validationError} (file on disk left untouched)");
                    }
                    else
                    {
                        var versionId = GitSupport.Snapshot(_filePath, "SpreadsheetTool auto-save");
                        _dirty = false;
                        Log.LogStep(versionId != null
                            ? $"SpreadsheetTool.Dispose: auto-saved '{_filePath}' (version '{versionId}')"
                            : $"SpreadsheetTool.Dispose: auto-saved '{_filePath}' (no changes)");
                    }
                }
            }
            catch (Exception ex)
            {
                Log.LogStep($"SpreadsheetTool.Dispose: auto-save failed — {ex.Message}");
            }
            _workbook?.Dispose();
        }

        /// <summary>
        /// Gets the current file path of this workbook, or null if not loaded.
        /// </summary>
        public string? FilePath => string.IsNullOrEmpty(_filePath) ? null : _filePath;

        // ──────────────────────────────────────────────
        //  Worksheet operations
        // ──────────────────────────────────────────────

        /// <summary>
        /// Lists all worksheet names in the workbook, in sheet order.
        /// </summary>
        /// <returns>Array of sheet names.</returns>
        public string[] GetSheetNames()
        {
            var names = new string[_workbook.Worksheets.Count];
            for (int i = 0; i < names.Length; i++)
                names[i] = _workbook.Worksheets[i].Name;
            Log.LogStep($"SpreadsheetTool.GetSheetNames: [{string.Join(", ", names)}]");
            return names;
        }

        /// <summary>
        /// Renames an existing worksheet.
        /// </summary>
        /// <param name="currentName">Current worksheet name (case-sensitive, from GetSheetNames()).</param>
        /// <param name="newName">New worksheet name.</param>
        /// <returns>"true", or "Error: …" when the sheet is missing or the name is taken.</returns>
        public string RenameWorksheet(string currentName, string newName)
        {
            if (string.IsNullOrWhiteSpace(newName)) return "Error: new worksheet name is required";
            var ws = FindSheet(currentName);
            if (ws == null) return $"Error: worksheet '{currentName}' not found";
            if (FindSheet(newName) != null) return $"Error: worksheet '{newName}' already exists";
            // Keep the per-sheet tracking with the renamed sheet.
            foreach (var map in new[] { _touchedCols, _bestFitCols })
                if (map.TryGetValue(currentName, out var set))
                {
                    map.Remove(currentName);
                    map[newName] = set;
                }
            ws.Name = newName;
            Log.LogStep($"SpreadsheetTool.RenameWorksheet: '{currentName}' → '{newName}'");
            _dirty = true;
            return "true";
        }

        /// <summary>
        /// Shows or hides gridlines on a worksheet.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="show">True to show gridlines, false to hide them.</param>
        /// <returns>"true", or "Error: …" when the sheet is missing.</returns>
        public string ShowGridlines(string sheetName, bool show)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            ws.ShowGridlines = show;
            Log.LogStep($"SpreadsheetTool.ShowGridlines: '{sheetName}' show={show}");
            _dirty = true;
            return "true";
        }

        /// <summary>
        /// Shows or hides row and column headers (the gray 1,2,3... / A,B,C... area) on a worksheet.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="show">True to show headers, false to hide them.</param>
        /// <returns>"true", or "Error: …" when the sheet is missing.</returns>
        public string ShowRowColumnHeaders(string sheetName, bool show)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            ws.ShowRowColumnHeaders = show;
            Log.LogStep($"SpreadsheetTool.ShowRowColumnHeaders: '{sheetName}' show={show}");
            _dirty = true;
            return "true";
        }

        /// <summary>
        /// Sets the zoom percentage for a worksheet (10-400).
        /// 100 = normal zoom.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="zoomPercentage">Zoom level between 10 and 400.</param>
        /// <returns>"true", or "Error: …" when the sheet is missing or the zoom is out of range.</returns>
        public string SetZoom(string sheetName, int zoomPercentage)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            if (zoomPercentage < 10 || zoomPercentage > 400) return $"Error: zoom must be between 10 and 400 (got {zoomPercentage})";
            ws.Zoom = zoomPercentage;
            Log.LogStep($"SpreadsheetTool.SetZoom: '{sheetName}' zoom={zoomPercentage}");
            _dirty = true;
            return "true";
        }

        /// <summary>
        /// Protects a worksheet so its structure cannot be modified.
        /// After protection, cells marked as locked (true by default) become read-only.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <returns>"true", or "Error: …" when the sheet is missing.</returns>
        public string ProtectSheet(string sheetName)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            ws.Protect();
            Log.LogStep($"SpreadsheetTool.ProtectSheet: '{sheetName}'");
            _dirty = true;
            return "true";
        }

        /// <summary>
        /// Removes protection from a previously protected worksheet,
        /// allowing edits to locked cells again.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <returns>"true", or "Error: …" when the sheet is missing.</returns>
        public string UnprotectSheet(string sheetName)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            ws.Unprotect();
            Log.LogStep($"SpreadsheetTool.UnprotectSheet: '{sheetName}'");
            _dirty = true;
            return "true";
        }

        /// <summary>Adds a new, empty worksheet to the workbook.</summary>
        /// <param name="name">Name for the new worksheet (e.g. "Riepilogo"). Must be unique in the workbook.</param>
        /// <returns>"true", or "Error: …" when the name is invalid or already taken.</returns>
        public string AddWorksheet(string name)
        {
            try
            {
                _workbook.Worksheets.Add(name);
                _dirty = true;
                Log.LogStep($"SpreadsheetTool.AddWorksheet: added '{name}'");
                return "true";
            }
            catch (Exception ex)
            {
                Log.LogStep($"SpreadsheetTool.AddWorksheet: FAILED — {ex.Message}");
                return $"Error: {ex.Message}";
            }
        }

        // ──────────────────────────────────────────────
        //  Cell read / write
        // ──────────────────────────────────────────────

        /// <summary>
        /// Returns the display string value of a cell.
        /// Applies number/date formatting if the cell has any.
        /// </summary>
        /// <param name="sheetName">Worksheet name (case-sensitive, from GetSheetNames()).</param>
        /// <param name="cellReference">Cell reference in A1 notation (e.g. "A1", "C5", "AB12").</param>
        /// <returns>The cell's display value, or null if not found.</returns>
        public string? GetCellValue(string sheetName, string cellReference)
        {
            var ws = FindSheet(sheetName);
            if (ws == null || !TryParseCellRef(cellReference).Ok) return null;
            return ws.Cells[cellReference]?.DisplayStringValue;
        }

        /// <summary>
        /// Sets the value of a cell. Auto-detects numbers, booleans, dates, and strings.
        /// </summary>
        /// <param name="sheetName">Worksheet name (case-sensitive, from GetSheetNames()).</param>
        /// <param name="cellReference">Cell reference in A1 notation (e.g. "A1", "C5").</param>
        /// <param name="value">The value to write. Parsed automatically.</param>
        /// <returns>The area receipt for the written cell — sheet name, exact cell reference and
        /// the stored value (feedback for the agent to verify its own work), or "Error: …" when
        /// the sheet or the cell reference is invalid.</returns>
        public object SetCellValue(string sheetName, string cellReference, string value)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            var p = TryParseCellRef(cellReference);
            if (!p.Ok) return $"Error: {p.Error}";
            MarkColumnTouched(ws.Name, p.Col);
            SetCellValueAuto(ws.Cells[cellReference], value);
            _dirty = true;
            Log.LogStep($"SpreadsheetTool.SetCellValue: '{sheetName}'!{cellReference} = '{value}'");
            return DescribeArea(ws, p.Row, p.Col, p.Row, p.Col);
        }

        /// <summary>
        /// Gets the formula of a cell (e.g. "=SUM(A1:A10)").
        /// Returns null if the cell has no formula.
        /// Note: formulas are stored and round-tripped but are NOT recalculated automatically.
        /// </summary>
        /// <param name="sheetName">Worksheet name (case-sensitive, from GetSheetNames()).</param>
        /// <param name="cellReference">Cell reference in A1 notation (e.g. "A1").</param>
        /// <returns>The formula with leading '=', or null.</returns>
        public string? GetCellFormula(string sheetName, string cellReference)
        {
            var ws = FindSheet(sheetName);
            if (ws == null || !TryParseCellRef(cellReference).Ok) return null;
            var f = ws.Cells[cellReference]?.Formula;
            return string.IsNullOrEmpty(f) ? null : f;
        }

        /// <summary>
        /// Sets a formula on a cell (e.g. "=SUM(A1:A10)").
        /// Formulas are stored and round-tripped but not recalculated automatically.
        /// </summary>
        /// <param name="sheetName">Worksheet name (case-sensitive, from GetSheetNames()).</param>
        /// <param name="cellReference">Cell reference in A1 notation (e.g. "A1").</param>
        /// <param name="formula">Formula including the leading '=' (e.g. "=SUM(B2:B10)").</param>
        /// <returns>The area receipt for the written cell (sheet, cell reference, formula —
        /// feedback for the agent to verify its own work), or "Error: …" on invalid input.</returns>
        public object SetCellFormula(string sheetName, string cellReference, string formula)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            var p = TryParseCellRef(cellReference);
            if (!p.Ok) return $"Error: {p.Error}";
            var cell = ws.Cells[cellReference];
            if (cell == null) return $"Error: cell '{cellReference}' not found";
            MarkColumnTouched(ws.Name, p.Col);
            cell.Formula = formula;
            Log.LogStep($"SpreadsheetTool.SetCellFormula: '{sheetName}'!{cellReference} = {formula}");
            _dirty = true;
            return DescribeArea(ws, p.Row, p.Col, p.Row, p.Col);
        }

        /// <summary>
        /// Returns value, formula and type of a cell in a single call — the agent needs one
        /// round-trip instead of GetCellValue + GetCellFormula + GetCellType.
        /// Value is the display string; Formula is null when the cell has none.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="cellReference">Cell reference in A1 notation (e.g. "A1").</param>
        /// <returns>JSON with "value", "formula" and "type" keys, or null if not found.</returns>
        public string? GetCellInfo(string sheetName, string cellReference)
        {
            var ws = FindSheet(sheetName);
            if (ws == null || !TryParseCellRef(cellReference).Ok) return null;
            var cell = ws.Cells[cellReference];
            if (cell == null) return null;
            var f = cell.Formula;
            return System.Text.Json.JsonSerializer.Serialize(new Dictionary<string, object?>
            {
                ["value"] = cell.DisplayStringValue,
                ["formula"] = string.IsNullOrEmpty(f) ? null : f,
                ["type"] = GetCellType(sheetName, cellReference),
            });
        }

        /// <summary>
        /// Returns the underlying value type of a cell.
        /// Possible values: "Unknown", "Null", "Numeric", "DateTime", "String", "Bool", "Error".
        /// Helps the agent determine how to interpret cell data before reading it.
        /// </summary>
        /// <param name="sheetName">Worksheet name (case-sensitive, from GetSheetNames()).</param>
        /// <param name="cellReference">Cell reference in A1 notation (e.g. "A1").</param>
        /// <returns>The value type name, or null if the cell is not found.</returns>
        public string? GetCellType(string sheetName, string cellReference)
        {
            var ws = FindSheet(sheetName);
            if (ws == null || !TryParseCellRef(cellReference).Ok) return null;
            var cell = ws.Cells[cellReference];
            if (cell == null) return null;
            return cell.Type switch
            {
                CellValueType.IsNull => "Null",
                CellValueType.IsNumeric => "Numeric",
                CellValueType.IsDateTime => "DateTime",
                CellValueType.IsString => "String",
                CellValueType.IsBool => "Bool",
                CellValueType.IsError => "Error",
                _ => "Unknown",
            };
        }

        // ──────────────────────────────────────────────
        //  Bulk operations
        // ──────────────────────────────────────────────

        /// <summary>
        /// Reads a rectangular range as a 2D string array.
        /// Empty cells are returned as empty strings.
        /// </summary>
        /// <param name="sheetName">Worksheet name (case-sensitive, from GetSheetNames()).</param>
        /// <param name="startCell">Top-left cell (e.g. "A1").</param>
        /// <param name="endCell">Bottom-right cell (e.g. "C10").</param>
        /// <param name="detailed">When true, returns an enriched JSON object instead of the raw 2D
        /// array: the sheet name and the exact A1 range (position on the page), plus SPARSE
        /// "formulas"/"types"/"formats" objects that list ONLY the cells carrying that information
        /// (cells not listed in "types" are numeric) — so the agent gets a verifiable, compact
        /// picture of the area as feedback/memory for its work.</param>
        /// <returns>2D string array [row][col] (or the enriched object with detailed=true),
        /// or null if the sheet is not found.</returns>
        public object? GetRange(string sheetName, string startCell, string endCell, bool detailed = false)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return null;
            var a = TryParseCellRef(startCell);
            var b = TryParseCellRef(endCell);
            if (!a.Ok || !b.Ok) return null;

            var (startRow, startCol) = (Math.Min(a.Row, b.Row), Math.Min(a.Col, b.Col));
            var (endRow, endCol) = (Math.Max(a.Row, b.Row), Math.Max(a.Col, b.Col));

            if (detailed)
                return DescribeArea(ws, startRow, startCol, endRow, endCol);

            int rows = endRow - startRow + 1;
            int cols = endCol - startCol + 1;
            var result = new string[rows][];

            for (int r = 0; r < rows; r++)
            {
                result[r] = new string[cols];
                for (int c = 0; c < cols; c++)
                {
                    var cell = ws.Cells[startRow + r, startCol + c];
                    result[r][c] = cell?.DisplayStringValue ?? string.Empty;
                }
            }
            return result;
        }

        /// <summary>Builds the compact, SPARSE JSON representation of a rectangular area: sheet
        /// name + exact A1 range (position on the page) + dimensions + display values; the
        /// "formulas"/"types"/"formats" objects include ONLY the cells that carry that information
        /// (empty sections are omitted entirely — no null/empty fields — to keep the token cost
        /// low). Cells not listed in "types" are numeric. Used by GetRange(detailed=true) and by
        /// every write method as the feedback/receipt of what was modified.</summary>
        private Dictionary<string, object?> DescribeArea(Worksheet ws, int startRow, int startCol, int endRow, int endCol)
        {
            var values = new List<object[]>();
            var formulas = new Dictionary<string, object>();
            var types = new Dictionary<string, object>();
            var formats = new Dictionary<string, object>();
            for (int r = startRow; r <= endRow; r++)
            {
                var row = new List<object>();
                for (int c = startCol; c <= endCol; c++)
                {
                    var cell = ws.Cells[r, c];
                    row.Add(cell?.DisplayStringValue ?? "");
                    if (cell == null) continue;
                    var refName = CellRefFromIdx(r, c);
                    if (!string.IsNullOrEmpty(cell.Formula))
                        formulas[refName] = cell.Formula.StartsWith('=') ? cell.Formula : "=" + cell.Formula;
                    else if (cell.Type is CellValueType.IsString or CellValueType.IsDateTime
                        or CellValueType.IsBool or CellValueType.IsError)
                        types[refName] = cell.Type switch
                        {
                            CellValueType.IsString => "text",
                            CellValueType.IsDateTime => "date",
                            CellValueType.IsBool => "bool",
                            _ => "error",
                        };
                    var custom = cell.GetStyle().Custom;
                    if (!string.IsNullOrEmpty(custom))
                        formats[refName] = custom;
                }
                values.Add(row.ToArray());
            }
            var result = new Dictionary<string, object?>
            {
                ["sheet"] = ws.Name,
                ["range"] = $"{CellRefFromIdx(startRow, startCol)}:{CellRefFromIdx(endRow, endCol)}",
                ["rows"] = endRow - startRow + 1,
                ["columns"] = endCol - startCol + 1,
                ["values"] = values,
            };
            if (formulas.Count > 0) result["formulas"] = formulas;
            if (types.Count > 0) result["types"] = types;
            if (formats.Count > 0) result["formats"] = formats;
            return result;
        }

        /// <summary>
        /// Writes a 2D string array starting at the specified cell.
        /// Auto-detects number, boolean, date, and string values.
        /// </summary>
        /// <param name="sheetName">Worksheet name (case-sensitive, from GetSheetNames()).</param>
        /// <param name="startCell">Top-left cell (e.g. "A1").</param>
        /// <param name="values">2D array of STRINGS [row][col]; pass numbers as strings (e.g. "14500") — the tool auto-detects and stores them as numbers. Rows may have different lengths.</param>
        /// <returns>The area receipt for the written block (sheet, exact A1 range, dimensions,
        /// values, formulas/types/formats when present — feedback for the agent to verify its own
        /// work), or "Error: …" when the sheet/cell reference is invalid or the block is too
        /// large (nothing is written on error).</returns>
        public object SetRange(string sheetName, string startCell, string[][] values)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            if (values == null || values.Length == 0) return "Error: no rows to write";
            var p = TryParseCellRef(startCell);
            if (!p.Ok) return $"Error: {p.Error}";

            // Guard the whole block BEFORE writing a single cell: a partial write followed by a
            // later failure would leave the workbook half-modified with no way to report it.
            // Null rows are treated as empty rows — an irregular block must never throw.
            int maxCols = values.Max(r => r?.Length ?? 0);
            long area = values.Sum(r => (long)(r?.Length ?? 0));
            if (maxCols == 0) return "Error: no values to write";
            if (area > MaxCellArea) return $"Error: the values block is too large ({area} cells, max {MaxCellArea})";
            if (p.Row + values.Length > 1_048_576) return "Error: the values block extends past the last row (1048576)";

            var (startRow, startCol) = (p.Row, p.Col);

            for (int c = 0; c < maxCols; c++)
                MarkColumnTouched(ws.Name, startCol + c);

            for (int r = 0; r < values.Length; r++)
            {
                if (values[r] == null) continue;
                for (int c = 0; c < values[r].Length; c++)
                {
                    var cell = ws.Cells[startRow + r, startCol + c];
                    if (cell != null)
                        SetCellValueAuto(cell, values[r][c]);
                }
            }

            _dirty = true;
            Log.LogStep($"SpreadsheetTool.SetRange: '{sheetName}'!{startCell} ({values.Length} rows)");
            return DescribeArea(ws, startRow, startCol, startRow + values.Length - 1, startCol + maxCols - 1);
        }

        /// <summary>
        /// Appends rows after the last used row on the worksheet.
        /// The last used row is detected across ALL columns (values AND formulas), so rows are
        /// never appended on top of existing content.
        /// </summary>
        /// <param name="sheetName">Worksheet name (case-sensitive, from GetSheetNames()).</param>
        /// <param name="rows">Array of rows to append.</param>
        /// <returns>The area receipt for the appended block (sheet, exact A1 range, dimensions,
        /// values — feedback for the agent to verify its own work), or "Error: …" when the sheet
        /// is missing or the block is too large (nothing is written on error).</returns>
        public object AppendRows(string sheetName, string[][] rows)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            if (rows == null || rows.Length == 0 || rows[0].Length == 0) return "Error: no rows to append";

            // Append AFTER the true last used row: the max over ALL columns of a value-OR-formula
            // scan. A display-only scan on column A would (a) ignore data sitting in other columns
            // and (b) stop at the first formula cell (formulas display as empty without
            // recalculation) — both would append on top of existing content and overwrite it.
            int startRow = FindLastUsedRow(ws) + 1;

            int maxCols = rows.Max(r => r?.Length ?? 0);
            long area = rows.Sum(r => (long)(r?.Length ?? 0));
            if (maxCols == 0) return "Error: no values to append";
            if (area > MaxCellArea) return $"Error: the rows block is too large ({area} cells, max {MaxCellArea})";
            if (startRow + rows.Length > 1_048_576) return "Error: the rows block extends past the last row (1048576)";

            for (int c = 0; c < maxCols; c++)
                MarkColumnTouched(ws.Name, c);

            for (int r = 0; r < rows.Length; r++)
            {
                if (rows[r] == null) continue;
                for (int c = 0; c < rows[r].Length; c++)
                {
                    var cell = ws.Cells[startRow + r, c];
                    if (cell != null)
                        SetCellValueAuto(cell, rows[r][c]);
                }
            }

            _dirty = true;
            Log.LogStep($"SpreadsheetTool.AppendRows: '{sheetName}' ({rows.Length} rows from row {startRow})");
            return DescribeArea(ws, startRow, 0, startRow + rows.Length - 1, maxCols - 1);
        }

        // ──────────────────────────────────────────────
        //  Merge
        // ──────────────────────────────────────────────

        /// <summary>
        /// Merges a rectangular range of cells into one cell.
        /// Only the upper-left cell value is preserved.
        /// </summary>
        /// <param name="sheetName">Worksheet name (case-sensitive, from GetSheetNames()).</param>
        /// <param name="range">Range in A1 notation (e.g. "A1:C3").</param>
        /// <returns>"true", or "Error: …" when the sheet or the range is invalid.</returns>
        public string MergeCells(string sheetName, string range)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            var p = TryParseRange(range);
            if (!p.Ok) return $"Error: {p.Error}";

            ws.Cells.Merge(p.R1, p.C1, p.R2 - p.R1 + 1, p.C2 - p.C1 + 1);
            _dirty = true;
            Log.LogStep($"SpreadsheetTool.MergeCells: '{sheetName}'!{range}");
            return "true";
        }

        // ──────────────────────────────────────────────
        //  Style — all-in-one
        // ──────────────────────────────────────────────

        /// <summary>
        /// Applies multiple style properties in a single call — the ONLY style method the agent
        /// needs (font, fill, alignment, wrap, number format and borders).
        /// Only non-null/non-default parameters are applied.
        /// Colors: "#RRGGBB". Pass fillColorHex "none" to remove the fill.
        /// Horizontal: "Left","Center","Right". Vertical: "Top","Center","Bottom".
        /// Border styles: "Thin","Medium","Thick","Dotted","Dashed","Double","Hair".
        /// Border sides: "All","Outline","Inside","Top","Bottom","Left","Right".
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="cellOrRange">Cell or range (e.g. "A1" or "A1:C10").</param>
        /// <param name="fontName">Font family.</param>
        /// <param name="fontSize">Font size in points.</param>
        /// <param name="bold">True = bold.</param>
        /// <param name="italic">True = italic.</param>
        /// <param name="fontColorHex">Font color "#RRGGBB".</param>
        /// <param name="fillColorHex">Background color "#RRGGBB", or "none" to remove the fill.</param>
        /// <param name="horizontalAlignment">Horizontal: "Left","Center","Right".</param>
        /// <param name="verticalAlignment">Vertical: "Top","Center","Bottom".</param>
        /// <param name="wrapText">True = wrap text.</param>
        /// <param name="numberFormat">Custom number format (e.g. "#,##0.00"). Pass it as a plain
        /// string WITHOUT surrounding quotes — a quoted string is a literal in Excel format syntax.</param>
        /// <param name="borderStyle">Border line style (e.g. "Thin"). Only applied when set.</param>
        /// <param name="borderSide">Which sides to border (default "All").</param>
        /// <param name="borderColorHex">Border color "#RRGGBB" (default black).</param>
        /// <returns>The style receipt — sheet, styled range and the list of parameters actually
        /// applied (feedback for the agent to verify its own work), or "Error: …" on invalid input.</returns>
        public object ApplyStyle(string sheetName, string cellOrRange,
            string? fontName = null, double fontSize = 0,
            bool? bold = null, bool? italic = null,
            string? fontColorHex = null, string? fillColorHex = null,
            string? horizontalAlignment = null, string? verticalAlignment = null,
            bool? wrapText = null, string? numberFormat = null,
            string? borderStyle = null, string? borderSide = null, string? borderColorHex = null)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            var guard = ValidateCellOrRange(cellOrRange);
            if (guard != null) return $"Error: {guard}";
            var p = TryParseRange(cellOrRange);
            if (!p.Ok) return $"Error: {p.Error}";
            var (r1, c1, r2, c2) = (p.R1, p.C1, p.R2, p.C2);

            for (int c = c1; c <= c2; c++)
                MarkColumnTouched(ws.Name, c);

            // Border parameters are resolved once, outside the cell loop. "Inside" borders only
            // the edges BETWEEN cells (outer perimeter excluded); "Outline" borders only the
            // outer perimeter; a single-cell range has no inner borders.
            BorderStyleType? border = string.IsNullOrEmpty(borderStyle) ? null : ParseBorderStyle(borderStyle);
            bool inside = borderSide?.Equals("Inside", StringComparison.OrdinalIgnoreCase) == true;
            bool outline = borderSide?.Equals("Outline", StringComparison.OrdinalIgnoreCase) == true;
            var sides = inside || outline
                ? new[] { "Top", "Bottom", "Left", "Right" }
                : ParseBorderSides(borderSide ?? "All");
            var borderColor = !string.IsNullOrEmpty(borderColorHex)
                ? ParseColor(borderColorHex)
                : Color.FromArgb(255, 0, 0, 0);

            for (int r = r1; r <= r2; r++)
            {
                for (int c = c1; c <= c2; c++)
                {
                    var cell = ws.Cells[r, c];
                    if (cell == null) continue;
                    var style = cell.GetStyle();

                    if (!string.IsNullOrEmpty(fontName)) style.Font.Name = fontName;
                    if (fontSize > 0) style.Font.Size = fontSize;
                    if (bold.HasValue) style.Font.IsBold = bold.Value;
                    if (italic.HasValue) style.Font.IsItalic = italic.Value;
                    if (!string.IsNullOrEmpty(fontColorHex)) style.Font.Color = ParseColor(fontColorHex);

                    if (!string.IsNullOrEmpty(fillColorHex))
                    {
                        if (fillColorHex.Equals("none", StringComparison.OrdinalIgnoreCase))
                            style.Pattern = FillPattern.None;
                        else
                        {
                            style.Pattern = FillPattern.Solid;
                            style.ForegroundColor = ParseColor(fillColorHex);
                        }
                    }

                    if (!string.IsNullOrEmpty(horizontalAlignment))
                        style.HorizontalAlignment = ParseHorizontalAlignment(horizontalAlignment);
                    if (!string.IsNullOrEmpty(verticalAlignment))
                        style.VerticalAlignment = ParseVerticalAlignment(verticalAlignment);
                    if (wrapText.HasValue)
                        style.WrapText = wrapText.Value;
                    if (!string.IsNullOrEmpty(numberFormat))
                        style.Custom = NormalizeNumberFormat(numberFormat);

                    if (border.HasValue)
                    {
                        var borders = style.Borders;
                        foreach (var side in sides)
                        {
                            if (inside && !IsInnerEdge(side, r, c, r1, c1, r2, c2)) continue;
                            if (outline && !IsOuterEdge(side, r, c, r1, c1, r2, c2)) continue;
                            switch (side)
                            {
                                case "Top":
                                    borders.Top.LineStyle = border.Value;
                                    borders.Top.Color = borderColor;
                                    break;
                                case "Bottom":
                                    borders.Bottom.LineStyle = border.Value;
                                    borders.Bottom.Color = borderColor;
                                    break;
                                case "Left":
                                    borders.Left.LineStyle = border.Value;
                                    borders.Left.Color = borderColor;
                                    break;
                                case "Right":
                                    borders.Right.LineStyle = border.Value;
                                    borders.Right.Color = borderColor;
                                    break;
                            }
                        }
                    }

                    cell.SetStyle(style);
                }
            }

            var applied = new List<object>();
            if (!string.IsNullOrEmpty(fontName)) applied.Add("font '" + fontName + "'");
            if (fontSize > 0) applied.Add("fontSize " + fontSize);
            if (bold.HasValue) applied.Add("bold=" + bold.Value.ToString().ToLowerInvariant());
            if (italic.HasValue) applied.Add("italic=" + italic.Value.ToString().ToLowerInvariant());
            if (!string.IsNullOrEmpty(fontColorHex)) applied.Add("fontColor " + fontColorHex);
            if (!string.IsNullOrEmpty(fillColorHex))
                applied.Add(fillColorHex.Equals("none", StringComparison.OrdinalIgnoreCase) ? "fill none" : "fill " + fillColorHex);
            if (!string.IsNullOrEmpty(horizontalAlignment)) applied.Add("hAlign " + horizontalAlignment);
            if (!string.IsNullOrEmpty(verticalAlignment)) applied.Add("vAlign " + verticalAlignment);
            if (wrapText.HasValue) applied.Add("wrapText=" + wrapText.Value.ToString().ToLowerInvariant());
            if (!string.IsNullOrEmpty(numberFormat)) applied.Add("numberFormat '" + NormalizeNumberFormat(numberFormat) + "'");
            if (border.HasValue) applied.Add("border " + borderStyle + (borderSide != null && borderSide != "All" ? " (" + borderSide + ")" : ""));

            Log.LogStep($"SpreadsheetTool.ApplyStyle: '{sheetName}'!{cellOrRange}");
            _dirty = true;
            return new Dictionary<string, object?>
            {
                ["sheet"] = ws.Name,
                ["range"] = $"{CellRefFromIdx(r1, c1)}:{CellRefFromIdx(r2, c2)}",
                ["applied"] = applied,
            };
        }

        /// <summary>True when the given side of cell (r,c) is an edge INSIDE the range (both
        /// neighbors in range) — used for "Inside" borders, which skip the outer perimeter.</summary>
        private static bool IsInnerEdge(string side, int r, int c, int r1, int c1, int r2, int c2) =>
            (side == "Top" && r > r1) || (side == "Bottom" && r < r2)
            || (side == "Left" && c > c1) || (side == "Right" && c < c2);

        /// <summary>True when the given side of cell (r,c) lies on the outer perimeter of the
        /// range — used for "Outline" borders.</summary>
        private static bool IsOuterEdge(string side, int r, int c, int r1, int c1, int r2, int c2) =>
            (side == "Top" && r == r1) || (side == "Bottom" && r == r2)
            || (side == "Left" && c == c1) || (side == "Right" && c == c2);

        // ──────────────────────────────────────────────
        //  Style — header row shortcut
        // ──────────────────────────────────────────────

        /// <summary>
        /// Applies a bold white-on-blue header style to the first row.
        /// Detects the used column count from row 0.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <returns>The style receipt — sheet, the formatted header range and the list of applied
        /// parameters (feedback for the agent to verify its own work), or "Error: …" when the
        /// sheet is missing or the header row is empty.</returns>
        public object FormatHeaderRow(string sheetName)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";

            int maxCol = 0;
            while (true)
            {
                var cell = ws.Cells[0, maxCol];
                if (cell == null || string.IsNullOrEmpty(cell.DisplayStringValue))
                    break;
                maxCol++;
            }
            if (maxCol == 0) return "Error: the header row is empty (no cells to format)";

            for (int c = 0; c < maxCol; c++)
            {
                var cell = ws.Cells[0, c];
                if (cell == null) continue;
                MarkColumnTouched(ws.Name, c);
                var style = cell.GetStyle();
                style.Font.IsBold = true;
                style.Font.Color = Color.FromArgb(255, 255, 255, 255);
                style.Pattern = FillPattern.Solid;
                style.ForegroundColor = Color.FromArgb(255, 34, 120, 212);
                cell.SetStyle(style);
            }
            Log.LogStep($"SpreadsheetTool.FormatHeaderRow: '{sheetName}' ({maxCol} columns)");
            _dirty = true;
            return new Dictionary<string, object?>
            {
                ["sheet"] = ws.Name,
                ["range"] = $"{CellRefFromIdx(0, 0)}:{CellRefFromIdx(0, maxCol - 1)}",
                ["applied"] = new object[] { "bold", "fontColor white", "fill #2278D4" },
            };
        }

        // ──────────────────────────────────────────────
        //  Row & column
        // ──────────────────────────────────────────────

        /// <summary>
        /// Sets the height of a specific row (0-based index).
        /// Height is in points (default row height is ~15 points).
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="rowIndex">0-based row index.</param>
        /// <param name="heightInPoints">Row height in points. Use -1 to reset to default.</param>
        /// <returns>"true", or "Error: …" if the row height was set.</returns>
        public string SetRowHeight(string sheetName, int rowIndex, double heightInPoints)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            if (rowIndex < 0 || rowIndex > 1_048_575) return $"Error: row index {rowIndex} out of range (0..1048575)";
            ws.Cells.Rows[rowIndex].Height = heightInPoints >= 0 ? heightInPoints : null;
            Log.LogStep($"SpreadsheetTool.SetRowHeight: '{sheetName}' row={rowIndex} height={heightInPoints}");
            _dirty = true;
            return "true";
        }

        /// <summary>
        /// Hides a specific row (0-based index).
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="rowIndex">0-based row index.</param>
        /// <returns>"true", or "Error: …" if the row was hidden.</returns>
        public string HideRow(string sheetName, int rowIndex)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            if (rowIndex < 0 || rowIndex > 1_048_575) return $"Error: row index {rowIndex} out of range (0..1048575)";
            ws.Cells.Rows[rowIndex].IsHidden = true;
            Log.LogStep($"SpreadsheetTool.HideRow: '{sheetName}' row={rowIndex}");
            _dirty = true;
            return "true";
        }

        /// <summary>
        /// Unhides a previously hidden row.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="rowIndex">0-based row index.</param>
        /// <returns>"true", or "Error: …" if the row was unhidden.</returns>
        public string UnhideRow(string sheetName, int rowIndex)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            if (rowIndex < 0 || rowIndex > 1_048_575) return $"Error: row index {rowIndex} out of range (0..1048575)";
            ws.Cells.Rows[rowIndex].IsHidden = false;
            Log.LogStep($"SpreadsheetTool.UnhideRow: '{sheetName}' row={rowIndex}");
            _dirty = true;
            return "true";
        }

        /// <summary>
        /// Sets the width of a specific column (0-based index).
        /// Width is measured in characters of the default font.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="columnIndex">0-based column index.</param>
        /// <param name="widthInCharacters">Column width in character units.</param>
        /// <returns>"true", or "Error: …" if the column width was set.</returns>
        public string SetColumnWidth(string sheetName, int columnIndex, double widthInCharacters)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            if (columnIndex < 0 || columnIndex > 16_383) return $"Error: column index {columnIndex} out of range (0..16383)";
            ws.Cells.Columns[columnIndex].Width = widthInCharacters;
            Log.LogStep($"SpreadsheetTool.SetColumnWidth: '{sheetName}' col={columnIndex} width={widthInCharacters}");
            _dirty = true;
            return "true";
        }

        /// <summary>
        /// Hides a specific column (0-based index).
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="columnIndex">0-based column index.</param>
        /// <returns>"true", or "Error: …" if the column was hidden.</returns>
        public string HideColumn(string sheetName, int columnIndex)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            if (columnIndex < 0 || columnIndex > 16_383) return $"Error: column index {columnIndex} out of range (0..16383)";
            ws.Cells.Columns[columnIndex].IsHidden = true;
            Log.LogStep($"SpreadsheetTool.HideColumn: '{sheetName}' col={columnIndex}");
            _dirty = true;
            return "true";
        }

        /// <summary>
        /// Unhides a previously hidden column.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="columnIndex">0-based column index.</param>
        /// <returns>"true", or "Error: …" if the column was unhidden.</returns>
        public string UnhideColumn(string sheetName, int columnIndex)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            if (columnIndex < 0 || columnIndex > 16_383) return $"Error: column index {columnIndex} out of range (0..16383)";
            ws.Cells.Columns[columnIndex].IsHidden = false;
            Log.LogStep($"SpreadsheetTool.UnhideColumn: '{sheetName}' col={columnIndex}");
            _dirty = true;
            return "true";
        }

        // ──────────────────────────────────────────────
        //  Charts
        // ──────────────────────────────────────────────

        /// <summary>
        /// Adds a new chart to a worksheet and returns its index as a string.
        /// The chart is placed in the specified cell-anchored rectangle.
        /// The data range is normalized (sheet name defaulted to the target sheet, "$" removed).
        /// Data convention: with a 2D block (several rows AND columns) the first row holds the
        /// series names, the first column holds the category labels, and every other column
        /// becomes a series — the classic matrix layout. A 1D range (one row or one column)
        /// becomes a single series without categories. The chart is saved with its data
        /// embedded, so it renders even in apps that do not refresh chart data on open.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="chartType">Chart type: "Column", "Bar", "Line", "Pie", "Area", "Scatter", "Doughnut", "Radar", etc.</param>
        /// <param name="dataRange">Cell range for chart data (e.g. "Sheet1!$A$1:$B$5" or "A1:B5").</param>
        /// <param name="upperLeftRow">Zero-based row for upper-left anchor.</param>
        /// <param name="upperLeftColumn">Zero-based column for upper-left anchor.</param>
        /// <param name="lowerRightRow">Zero-based row for lower-right anchor.</param>
        /// <param name="lowerRightColumn">Zero-based column for lower-right anchor.</param>
        /// <returns>The zero-based chart index as a string (e.g. "0"), or "Error: …" on failure.</returns>
        public string AddChart(string sheetName, string chartType, string dataRange,
            int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            var type = ParseChartType(chartType);
            if (type == null)
                return $"Error: unknown chart type '{chartType}' (valid: Column, Bar, Line, Pie, Area, Scatter, Doughnut, Radar, Stock, 3D variants)";
            if (string.IsNullOrWhiteSpace(dataRange)) return "Error: a data range is required";
            if (upperLeftRow < 0 || upperLeftColumn < 0 || lowerRightRow < upperLeftRow || lowerRightColumn < upperLeftColumn
                || lowerRightRow > 1_048_575 || lowerRightColumn > 16_383)
                return "Error: invalid anchor rectangle (zero-based, ordered rows/columns within the sheet)";

            var range = NormalizeChartRange(dataRange, ws.Name);
            var cellsPart = range[(range.LastIndexOf('!') + 1)..];
            var p = TryParseRange(cellsPart);
            if (!p.Ok) return $"Error: {p.Error}";

            try
            {
                var idx = ws.Charts.Add(type.Value, range, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
                _dirty = true;
                Log.LogStep($"SpreadsheetTool.AddChart: '{sheetName}' {chartType} (index {idx}) range='{range}'");
                return idx.ToString();
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        /// <summary>
        /// Normalizes a chart data range: defaults the sheet name to the target sheet when
        /// missing and removes "$" anchors. The agent may pass "A1:B5", "Vendite!A1:B5" or
        /// "Vendite!$A$1:$B$5" — all are accepted.
        /// </summary>
        private static string NormalizeChartRange(string dataRange, string defaultSheet)
        {
            var r = (dataRange ?? "").Trim();
            var bang = r.LastIndexOf('!');
            var sheet = bang >= 0 ? r[..bang].Trim() : defaultSheet;
            var cells = bang >= 0 ? r[(bang + 1)..].Trim() : r.Trim('$').Trim();
            cells = cells.Replace("$", "");
            // Accept a sheet name the caller already quoted ("'Monthly Data'!A1:C7" — valid Excel
            // syntax) instead of double-quoting it below, which produces "''Monthly Data''!A1:C7".
            if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[^1] == '\'')
                sheet = sheet[1..^1].Trim();
            // Quote sheet names containing characters that break the range syntax.
            if (sheet.Any(ch => !char.IsLetterOrDigit(ch) && ch != '_'))
                sheet = "'" + sheet + "'";
            return $"{sheet}!{cells}";
        }

        /// <summary>
        /// Returns summary info about all charts on a worksheet as a 2D array.
        /// Columns: Index, Name, Type, Position.
        /// The chart count is GetChartsInfo().Length - 1 (the first row is the header).
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <returns>2D string array with chart info, or null if sheet not found.</returns>
        public string[][]? GetChartsInfo(string sheetName)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return null;
            var info = new List<string[]> { new[] { "Index", "Name", "Type", "UpperLeft", "LowerRight" } };
            for (int i = 0; i < ws.Charts.Count; i++)
            {
                var c = ws.Charts[i];
                info.Add(new[] {
                    i.ToString(),
                    c.Name,
                    c.ChartType.ToString(),
                    $"({c.UpperLeftRow},{c.UpperLeftColumn})",
                    $"({c.LowerRightRow},{c.LowerRightColumn})"
                });
            }
            return info.ToArray();
        }

        // ──────────────────────────────────────────────
        //  Hyperlinks
        // ──────────────────────────────────────────────

        /// <summary>
        /// Adds a hyperlink to a cell or range.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="cellName">Top-left cell reference (e.g. "A1").</param>
        /// <param name="totalRows">Number of rows the hyperlink spans.</param>
        /// <param name="totalColumns">Number of columns the hyperlink spans.</param>
        /// <param name="address">URL, file path, or email address.</param>
        /// <returns>"true", or "Error: …" when the sheet/cell is invalid.</returns>
        public string AddHyperlink(string sheetName, string cellName, int totalRows, int totalColumns, string address)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            var p = TryParseCellRef(cellName);
            if (!p.Ok) return $"Error: {p.Error}";
            if (totalRows < 1 || totalColumns < 1) return "Error: totalRows and totalColumns must be at least 1";
            ws.Hyperlinks.Add(cellName, totalRows, totalColumns, address);
            Log.LogStep($"SpreadsheetTool.AddHyperlink: '{sheetName}'!{cellName} → {address}");
            _dirty = true;
            return "true";
        }

        /// <summary>
        /// Returns all hyperlinks on a worksheet as a 2D array.
        /// Columns: Index, Area, Address, ScreenTip, TextToDisplay.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <returns>2D string array, or null if sheet not found.</returns>
        public string[][]? GetHyperlinks(string sheetName)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return null;
            var info = new List<string[]> { new[] { "Index", "Area", "Address", "ScreenTip", "TextToDisplay" } };
            for (int i = 0; i < ws.Hyperlinks.Count; i++)
            {
                var h = ws.Hyperlinks[i];
                info.Add(new[] { i.ToString(), h.Area, h.Address, h.ScreenTip ?? "", h.TextToDisplay ?? "" });
            }
            return info.ToArray();
        }

        /// <summary>
        /// Removes a hyperlink by its zero-based index.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="index">Zero-based hyperlink index (from GetHyperlinks()).</param>
        /// <returns>"true", or "Error: …" when the sheet or the index is invalid.</returns>
        public string RemoveHyperlink(string sheetName, int index)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            if (index < 0 || index >= ws.Hyperlinks.Count) return $"Error: hyperlink index {index} out of range (0..{ws.Hyperlinks.Count - 1})";
            ws.Hyperlinks.RemoveAt(index);
            Log.LogStep($"SpreadsheetTool.RemoveHyperlink: '{sheetName}' index={index}");
            _dirty = true;
            return "true";
        }

        // ──────────────────────────────────────────────
        //  Comments
        // ──────────────────────────────────────────────

        /// <summary>
        /// Adds a comment (note) to a cell.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="cellReference">Cell reference in A1 notation (e.g. "A1").</param>
        /// <param name="text">Comment text.</param>
        /// <param name="author">Optional author name.</param>
        /// <returns>"true", or "Error: …" when the sheet or the cell reference is invalid.</returns>
        public string AddComment(string sheetName, string cellReference, string text, string? author = null)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            if (!TryParseCellRef(cellReference).Ok) return $"Error: invalid cell reference '{cellReference}'";
            var comment = ws.Comments.Add(cellReference);
            comment.Note = text;
            if (!string.IsNullOrEmpty(author)) comment.Author = author;
            Log.LogStep($"SpreadsheetTool.AddComment: '{sheetName}'!{cellReference}");
            _dirty = true;
            return "true";
        }

        /// <summary>
        /// Gets the comment text at a specific cell, or null if no comment exists.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="cellReference">Cell reference (e.g. "A1").</param>
        /// <returns>The comment text, or null.</returns>
        public string? GetComment(string sheetName, string cellReference)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return null;
            var comment = ws.Comments[cellReference];
            return comment?.Note;
        }

        /// <summary>
        /// Lists all comments on a worksheet as a 2D array.
        /// Columns: Cell, Author, Note, Visible.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <returns>2D string array, or null if sheet not found.</returns>
        public string[][]? GetComments(string sheetName)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return null;
            var info = new List<string[]> { new[] { "Cell", "Author", "Note", "Visible" } };
            for (int i = 0; i < ws.Comments.Count; i++)
            {
                var c = ws.Comments[i];
                info.Add(new[] {
                    $"({c.Row},{c.Column})",
                    c.Author ?? "",
                    c.Note ?? "",
                    c.IsVisible ? "Yes" : "No"
                });
            }
            return info.ToArray();
        }

        /// <summary>
        /// Removes a comment from a cell.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="cellReference">Cell reference (e.g. "A1").</param>
        /// <returns>"true", or "Error: …" when the sheet or the cell reference is invalid.</returns>
        public string RemoveComment(string sheetName, string cellReference)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            if (!TryParseCellRef(cellReference).Ok) return $"Error: invalid cell reference '{cellReference}'";
            ws.Comments.RemoveAt(cellReference);
            Log.LogStep($"SpreadsheetTool.RemoveComment: '{sheetName}'!{cellReference}");
            _dirty = true;
            return "true";
        }

        // ──────────────────────────────────────────────
        //  Pictures
        // ──────────────────────────────────────────────

        /// <summary>
        /// Adds a picture to a worksheet from a file path.
        /// The picture is anchored to a cell rectangle.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="imageFilePath">Path to the image file, Unix style relative to the workspace root (e.g. "/folder/image.png").</param>
        /// <param name="upperLeftRow">Zero-based row for upper-left anchor.</param>
        /// <param name="upperLeftColumn">Zero-based column for upper-left anchor.</param>
        /// <param name="lowerRightRow">Zero-based row for lower-right anchor.</param>
        /// <param name="lowerRightColumn">Zero-based column for lower-right anchor.</param>
        /// <returns>"true", or "Error: …" when the sheet or the parameters are invalid.</returns>
        public string AddPicture(string sheetName, string imageFilePath,
            int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            if (string.IsNullOrWhiteSpace(imageFilePath)) return "Error: an image file path is required";
            if (upperLeftRow < 0 || upperLeftColumn < 0 || lowerRightRow < upperLeftRow || lowerRightColumn < upperLeftColumn)
                return "Error: invalid anchor rectangle (zero-based, ordered rows/columns)";
            try
            {
                ws.Pictures.Add(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, imageFilePath);
                Log.LogStep($"SpreadsheetTool.AddPicture: '{sheetName}' {imageFilePath}");
                _dirty = true;
                return "true";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        /// <summary>
        /// Removes a picture by its zero-based index.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="index">Zero-based picture index (0 to GetPicturesCount()-1).</param>
        /// <returns>"true", or "Error: …" when the sheet or the index is invalid.</returns>
        public string RemovePicture(string sheetName, int index)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            if (index < 0 || index >= ws.Pictures.Count) return $"Error: picture index {index} out of range (0..{ws.Pictures.Count - 1})";
            ws.Pictures.RemoveAt(index);
            Log.LogStep($"SpreadsheetTool.RemovePicture: '{sheetName}' index={index}");
            _dirty = true;
            return "true";
        }

        /// <summary>
        /// Returns the number of pictures on a worksheet.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <returns>Picture count, or -1 if not found.</returns>
        public int GetPicturesCount(string sheetName)
        {
            return FindSheet(sheetName)?.Pictures.Count ?? -1;
        }

        // ──────────────────────────────────────────────
        //  ListObjects (Excel Tables)
        // ──────────────────────────────────────────────

        /// <summary>Adds an Excel table (ListObject) covering the cell range on a worksheet.</summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="startCell">Top-left cell as a SEPARATE parameter, e.g. "A1" — not a combined range like "A1:D10".</param>
        /// <param name="endCell">Bottom-right cell as a SEPARATE parameter, e.g. "D55".</param>
        /// <param name="hasHeaders">True if the first row contains column headers.</param>
        /// <returns>"true", or "Error: …" when the sheet or the range is invalid.</returns>
        public string AddTable(string sheetName, string startCell, string endCell, bool hasHeaders = true)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            var a = TryParseCellRef(startCell);
            var b = TryParseCellRef(endCell);
            if (!a.Ok) return $"Error: {a.Error}";
            if (!b.Ok) return $"Error: {b.Error}";
            if (a.Row > b.Row || a.Col > b.Col) return "Error: startCell must be the top-left corner (start <= end)";
            ws.ListObjects.Add(startCell, endCell, hasHeaders);
            Log.LogStep($"SpreadsheetTool.AddTable: '{sheetName}'!{startCell}:{endCell}");
            _dirty = true;
            return "true";
        }

        /// <summary>
        /// Lists all tables on a worksheet as a 2D array.
        /// Columns: Index, Name, Range, HeaderRow, TotalsRow.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <returns>2D string array, or null if sheet not found.</returns>
        public string[][]? GetTables(string sheetName)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return null;
            var info = new List<string[]> { new[] { "Index", "Name", "Range", "HeaderRow", "TotalsRow" } };
            for (int i = 0; i < ws.ListObjects.Count; i++)
            {
                var t = ws.ListObjects[i];
                info.Add(new[] {
                    i.ToString(),
                    t.DisplayName,
                    $"({t.StartRow},{t.StartColumn})-({t.EndRow},{t.EndColumn})",
                    t.ShowHeaderRow ? "Yes" : "No",
                    t.ShowTotals ? "Yes" : "No"
                });
            }
            return info.ToArray();
        }

        /// <summary>
        /// Removes an Excel table by its zero-based index.
        /// The cell data is preserved.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="index">Zero-based table index (from GetTables()).</param>
        /// <returns>"true", or "Error: …" when the sheet or the index is invalid.</returns>
        public string RemoveTable(string sheetName, int index)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            if (index < 0 || index >= ws.ListObjects.Count) return $"Error: table index {index} out of range (0..{ws.ListObjects.Count - 1})";
            ws.ListObjects.RemoveAt(index);
            Log.LogStep($"SpreadsheetTool.RemoveTable: '{sheetName}' index={index}");
            _dirty = true;
            return "true";
        }

        // ──────────────────────────────────────────────
        //  AutoFilter
        // ──────────────────────────────────────────────

        /// <summary>
        /// Enables the AutoFilter (drop-down arrows) on a worksheet for the specified range.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="range">Range in A1 notation (e.g. "A1:C10").</param>
        /// <returns>"true", or "Error: …" when the sheet or the range is invalid.</returns>
        public string SetAutoFilter(string sheetName, string range)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            var p = TryParseRange(range);
            if (!p.Ok) return $"Error: {p.Error}";
            ws.AutoFilter.Range = range;
            Log.LogStep($"SpreadsheetTool.SetAutoFilter: '{sheetName}'!{range}");
            _dirty = true;
            return "true";
        }

        /// <summary>
        /// Removes the AutoFilter from a worksheet.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <returns>"true", or "Error: …" if the AutoFilter was removed.</returns>
        public string RemoveAutoFilter(string sheetName)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            ws.AutoFilter.Clear();
            Log.LogStep($"SpreadsheetTool.RemoveAutoFilter: '{sheetName}'");
            _dirty = true;
            return "true";
        }

        // ──────────────────────────────────────────────
        //  DefinedNames (named ranges / formula variables)
        // ──────────────────────────────────────────────

        /// <summary>
        /// Adds a defined name (named range / formula variable) to the workbook.
        /// Defined names can be used in cell formulas as variables.
        /// Example: AddDefinedName("TaxRate", "=0.22") → use "=A1 * TaxRate" in any cell.
        /// Example: AddDefinedName("SalesData", "=Sheet1!$A$1:$C$100") → use "=SUM(SalesData)".
        /// </summary>
        /// <param name="name">Variable name (e.g. "TaxRate", "SalesData"). Must be unique.</param>
        /// <param name="formula">Formula or reference the name points to (e.g. "=0.22", "=Sheet1!$A$1:$C$100").</param>
        /// <param name="localSheetName">Optional: if set, the name is scoped to this sheet only (from GetSheetNames()).</param>
        /// <returns>"true", or "Error: …" when the name/sheet is invalid or already taken.</returns>
        public string AddDefinedName(string name, string formula, string? localSheetName = null)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Error: a defined name is required";
            if (string.IsNullOrWhiteSpace(formula)) return "Error: a formula is required";
            try
            {
                int? sheetIndex = null;
                if (!string.IsNullOrEmpty(localSheetName))
                {
                    var ws = FindSheet(localSheetName);
                    if (ws == null) return $"Error: worksheet '{localSheetName}' not found";
                    // Find the sheet index
                    for (int i = 0; i < _workbook.Worksheets.Count; i++)
                    {
                        if (_workbook.Worksheets[i].Name == localSheetName)
                        { sheetIndex = i; break; }
                    }
                }
                _workbook.DefinedNames.Add(name, formula, sheetIndex);
                _dirty = true;
                Log.LogStep($"SpreadsheetTool.AddDefinedName: '{name}' = {formula}");
                return "true";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        /// <summary>
        /// Lists all defined names in the workbook as a 2D array.
        /// Columns: Name, Formula, Scope, Hidden, Comment.
        /// </summary>
        /// <returns>2D string array with defined name info.</returns>
        public string[][]? GetDefinedNames()
        {
            try
            {
                if (_workbook.DefinedNames.Count == 0)
                    return new[] { new[] { "Name", "Formula", "Scope", "Hidden", "Comment" } };

                var info = new List<string[]> { new[] { "Name", "Formula", "Scope", "Hidden", "Comment" } };
                for (int i = 0; i < _workbook.DefinedNames.Count; i++)
                {
                    var dn = _workbook.DefinedNames[i];
                    var scope = dn.LocalSheetIndex.HasValue
                        ? _workbook.Worksheets[dn.LocalSheetIndex.Value].Name
                        : "Workbook";
                    info.Add(new[] {
                        dn.Name,
                        dn.Formula,
                        scope,
                        dn.Hidden ? "Yes" : "No",
                        dn.Comment ?? ""
                    });
                }
                return info.ToArray();
            }
            catch { return null; }
        }

        /// <summary>
        /// Gets the formula that a defined name points to.
        /// Returns null if the name does not exist.
        /// </summary>
        /// <param name="name">Defined name (from GetDefinedNames()).</param>
        /// <returns>The formula string (e.g. "=0.22", "=Sheet1!$A$1:$C$100"), or null.</returns>
        public string? GetDefinedNameFormula(string name)
        {
            try
            {
                for (int i = 0; i < _workbook.DefinedNames.Count; i++)
                {
                    if (_workbook.DefinedNames[i].Name == name)
                        return _workbook.DefinedNames[i].Formula;
                }
                return null;
            }
            catch { return null; }
        }

        /// <summary>
        /// Removes a defined name by its name.
        /// </summary>
        /// <param name="name">The defined name to remove (from GetDefinedNames()).</param>
        /// <returns>"true", or "Error: …" when the name does not exist.</returns>
        public string RemoveDefinedName(string name)
        {
            try
            {
                for (int i = 0; i < _workbook.DefinedNames.Count; i++)
                {
                    if (_workbook.DefinedNames[i].Name == name)
                    {
                        _workbook.DefinedNames.RemoveAt(i);
                        Log.LogStep($"SpreadsheetTool.RemoveDefinedName: '{name}'");
                        _dirty = true;
                        return "true";
                    }
                }
                return $"Error: defined name '{name}' not found";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        // ──────────────────────────────────────────────
        //  PageSetup
        // ──────────────────────────────────────────────

        /// <summary>
        /// Sets page setup properties for a worksheet.
        /// Only non-null/non-default parameters are applied.
        /// Orientation: "Portrait" or "Landscape".
        /// PaperSize: "A4", "Letter", "A3", "Legal", etc.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="orientation">"Portrait" or "Landscape". Null to keep current.</param>
        /// <param name="paperSize">"A4", "Letter", "A3", "Legal", etc. Null to keep current.</param>
        /// <param name="scale">Print scaling percentage (10-400). Null to keep current.</param>
        /// <param name="fitToPagesWide">Fit print to this many pages wide. Null to keep current.</param>
        /// <param name="fitToPagesTall">Fit print to this many pages tall. Null to keep current.</param>
        /// <param name="printArea">Print area range (e.g. "A1:C10"). Null to keep current.</param>
        /// <param name="centerHorizontally">True to center horizontally on page.</param>
        /// <param name="centerVertically">True to center vertically on page.</param>
        /// <returns>The page-setup receipt — sheet and the list of settings actually applied
        /// (feedback for the agent to verify its own work, e.g. that A4/fit-to-page really took
        /// effect), or "Error: …" when the sheet is missing.</returns>
        public object SetPageSetup(string sheetName,
            string? orientation = null, string? paperSize = null,
            int? scale = null, int? fitToPagesWide = null, int? fitToPagesTall = null,
            string? printArea = null, bool? centerHorizontally = null, bool? centerVertically = null)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            var ps = ws.PageSetup;

            if (!string.IsNullOrEmpty(orientation))
                ps.Orientation = orientation.Equals("Landscape", StringComparison.OrdinalIgnoreCase)
                    ? PageOrientationType.Landscape : PageOrientationType.Portrait;

            if (!string.IsNullOrEmpty(paperSize))
                ps.PaperSize = ParsePaperSize(paperSize);

            if (scale.HasValue) ps.Scale = scale.Value;
            if (fitToPagesWide.HasValue) ps.FitToPagesWide = fitToPagesWide.Value;
            if (fitToPagesTall.HasValue) ps.FitToPagesTall = fitToPagesTall.Value;
            if (!string.IsNullOrEmpty(printArea)) ps.PrintArea = printArea;
            if (centerHorizontally.HasValue) ps.CenterHorizontally = centerHorizontally.Value;
            if (centerVertically.HasValue) ps.CenterVertically = centerVertically.Value;

            Log.LogStep($"SpreadsheetTool.SetPageSetup: '{sheetName}'");
            _dirty = true;

            var applied = new List<object>();
            if (!string.IsNullOrEmpty(orientation)) applied.Add("orientation " + (ps.Orientation == PageOrientationType.Landscape ? "Landscape" : "Portrait"));
            if (!string.IsNullOrEmpty(paperSize)) applied.Add("paperSize " + ps.PaperSize);
            if (scale.HasValue) applied.Add("scale " + scale);
            if (fitToPagesWide.HasValue) applied.Add("fitToPagesWide " + fitToPagesWide);
            if (fitToPagesTall.HasValue) applied.Add("fitToPagesTall " + fitToPagesTall);
            if (!string.IsNullOrEmpty(printArea)) applied.Add("printArea '" + printArea + "'");
            if (centerHorizontally.HasValue) applied.Add("centerHorizontally " + centerHorizontally.Value.ToString().ToLowerInvariant());
            if (centerVertically.HasValue) applied.Add("centerVertically " + centerVertically.Value.ToString().ToLowerInvariant());
            return new Dictionary<string, object?>
            {
                ["sheet"] = ws.Name,
                ["applied"] = applied,
            };
        }

        /// <summary>
        /// Returns the current PageSetup settings for a worksheet as a 2D array.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <returns>2D string array with page setup properties, or null.</returns>
        public string[][]? GetPageSetup(string sheetName)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return null;
            var ps = ws.PageSetup;
            return new[] {
                new[] { "Property", "Value" },
                new[] { "Orientation", ps.Orientation.ToString() },
                new[] { "PaperSize", ps.PaperSize.ToString() },
                new[] { "Scale", ps.Scale?.ToString() ?? "(default)" },
                new[] { "FitToPagesWide", ps.FitToPagesWide?.ToString() ?? "(default)" },
                new[] { "FitToPagesTall", ps.FitToPagesTall?.ToString() ?? "(default)" },
                new[] { "PrintArea", ps.PrintArea ?? "(none)" },
                new[] { "CenterHorizontally", ps.CenterHorizontally ? "Yes" : "No" },
                new[] { "CenterVertically", ps.CenterVertically ? "Yes" : "No" },
                new[] { "PrintGridlines", ps.PrintGridlines ? "Yes" : "No" },
                new[] { "LeftMargin", ps.LeftMargin.ToString("F2") + " cm" },
                new[] { "RightMargin", ps.RightMargin.ToString("F2") + " cm" },
                new[] { "TopMargin", ps.TopMargin.ToString("F2") + " cm" },
                new[] { "BottomMargin", ps.BottomMargin.ToString("F2") + " cm" },
            };
        }

        // ──────────────────────────────────────────────
        //  ConditionalFormatting
        // ──────────────────────────────────────────────

        /// <summary>
        /// Adds conditional formatting to a cell range.
        /// Condition types: "CellValue", "Formula", "AboveAverage", "Top10", "Unique", "Duplicate", etc.
        /// Operators (for CellValue): "Between", "NotBetween", "Equal", "NotEqual", "GreaterThan", "LessThan", "GreaterOrEqual", "LessOrEqual".
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="range">Cell range in A1 notation (e.g. "A1:C10").</param>
        /// <param name="conditionType">Condition type: "CellValue", "Formula", etc.</param>
        /// <param name="operatorType">Operator (required for CellValue): "Between", "Equal", "GreaterThan", etc.</param>
        /// <param name="formula1">First formula or value for the condition.</param>
        /// <param name="formula2">Second formula (required for "Between" / "NotBetween").</param>
        /// <param name="fontColorHex">Font color "#RRGGBB" when condition is met.</param>
        /// <param name="fillColorHex">Fill color "#RRGGBB" when condition is met.</param>
        /// <param name="bold">True for bold when condition is met.</param>
        /// <returns>"true", or "Error: …" when the sheet or the range is invalid.</returns>
        public string AddConditionalFormat(string sheetName, string range,
            string conditionType, string? operatorType = null,
            string? formula1 = null, string? formula2 = null,
            string? fontColorHex = null, string? fillColorHex = null, bool? bold = null)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return $"Error: worksheet '{sheetName}' not found";
            var p = TryParseRange(range);
            if (!p.Ok) return $"Error: {p.Error}";

            var area = new CellArea
            {
                StartRow = p.R1,
                StartColumn = p.C1,
                EndRow = p.R2,
                EndColumn = p.C2
            };

            var condType = ParseFormatConditionType(conditionType);
            var opType = operatorType != null ? ParseOperatorType(operatorType) : OperatorType.None;

            int ci = ws.ConditionalFormattings.Add();
            int fcIdx = ws.ConditionalFormattings[ci].AddCondition(condType, opType, formula1 ?? "", formula2 ?? "");
            ws.ConditionalFormattings[ci].AddArea(area);

            if (!string.IsNullOrEmpty(fontColorHex) || !string.IsNullOrEmpty(fillColorHex) || bold.HasValue)
            {
                var fc = ws.ConditionalFormattings[ci][fcIdx];
                var style = fc.Style;
                if (!string.IsNullOrEmpty(fontColorHex)) style.Font.Color = ParseColor(fontColorHex);
                if (!string.IsNullOrEmpty(fillColorHex)) { style.Pattern = FillPattern.Solid; style.ForegroundColor = ParseColor(fillColorHex); }
                if (bold.HasValue) style.Font.IsBold = bold.Value;
                fc.Style = style;
            }

            Log.LogStep($"SpreadsheetTool.AddConditionalFormat: '{sheetName}'!{range} ({conditionType})");
            _dirty = true;
            return "true";
        }

        // ──────────────────────────────────────────────
        //  Worksheet description (LLM-friendly JSON)
        // ──────────────────────────────────────────────

        /// <summary>
        /// Produces a compact but complete JSON description of one or all worksheets.
        /// Designed for LLM consumption: includes headers, sample data, formula patterns,
        /// tables, charts, comments, merged cells, AutoFilter, and defined names.
        /// Formula patterns detect repeating formulas across ranges and report the pattern
        /// rather than listing every cell — e.g. "=B{r}*D{r}" applied to "C2:C100".
        /// Use this method FIRST when you do not know the worksheet content, then zoom into
        /// the areas of interest with GetRange(detailed=true).
        /// </summary>
        /// <param name="sheetName">
        /// Specific worksheet to describe (from GetSheetNames()), or null to describe all worksheets.
        /// </param>
        /// <param name="sampleRowCount">
        /// Number of sample data rows to include per sheet (default 5, max 20).
        /// </param>
        /// <returns>A JSON string describing the worksheet structure and content.</returns>
        public string DescribeWorksheet(string? sheetName = null, int sampleRowCount = 5)
        {
            if (_workbook == null)
                return $"{{\"error\": \"No workbook loaded. Call {AIOrchestrator.Utility.ToSnakeCase(nameof(Open))}(filePath) first.\"}}";

            sampleRowCount = Math.Clamp(sampleRowCount, 1, 20);

            var result = new Dictionary<string, object?>
            {
                ["filePath"] = SandboxPath.ToAgent(_filePath),
                ["sheets"] = new List<object>()
            };

            int sheetCount = _workbook.Worksheets.Count;
            var sheetsList = (List<object>)result["sheets"]!;

            for (int si = 0; si < sheetCount; si++)
            {
                var ws = _workbook.Worksheets[si];
                if (sheetName != null && !ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    continue;

                var sheetDesc = DescribeOneSheet(ws, si, sampleRowCount);
                sheetsList.Add(sheetDesc);
            }

            // Add workbook-level defined names
            var dnList = new List<Dictionary<string, object?>>();
            for (int i = 0; i < _workbook.DefinedNames.Count; i++)
            {
                var dn = _workbook.DefinedNames[i];
                dnList.Add(new Dictionary<string, object?>
                {
                    ["name"] = dn.Name,
                    ["formula"] = dn.Formula,
                    ["scope"] = dn.LocalSheetIndex.HasValue
                        ? _workbook.Worksheets[dn.LocalSheetIndex.Value].Name
                        : "Workbook",
                    ["hidden"] = dn.Hidden
                });
            }
            if (dnList.Count > 0)
                result["definedNames"] = dnList;

            return System.Text.Json.JsonSerializer.Serialize(result, new System.Text.Json.JsonSerializerOptions
            {
                WriteIndented = false,
                DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
            });
        }

        private Dictionary<string, object?> DescribeOneSheet(Worksheet ws, int sheetIndex, int sampleRowCount)
        {
            var desc = new Dictionary<string, object?>
            {
                ["name"] = ws.Name,
                ["index"] = sheetIndex,
            };

            // Determine used range. Formula-only cells display as empty (no recalculation), so
            // the display scan must also count cells that carry a formula — otherwise a sheet
            // whose values are all computed (e.g. a KPI sheet) looks empty to the agent.
            int usedRows = 0, usedCols = 0;
            for (int r = 0; r < 5000; r++)
            {
                bool any = false;
                for (int c = 0; c < 500; c++)
                {
                    var cell = ws.Cells[r, c];
                    if (cell != null && (!string.IsNullOrEmpty(cell.DisplayStringValue) || !string.IsNullOrEmpty(cell.Formula)))
                    { any = true; if (c >= usedCols) usedCols = c + 1; }
                }
                if (any) usedRows = r + 1;
                else if (r > usedRows + 5) break; // 5 empty rows → stop
            }
            desc["usedRange"] = $"A1:{CellRefFromIdx(usedRows - 1, usedCols - 1)}";
            desc["usedRowCount"] = usedRows;
            desc["usedColumnCount"] = usedCols;

            // Gridlines, zoom, visibility
            desc["showGridlines"] = ws.ShowGridlines;
            desc["zoom"] = ws.Zoom;

            // Column info from first row (headers)
            var columns = new List<Dictionary<string, object?>>();
            for (int c = 0; c < usedCols; c++)
            {
                var colInfo = new Dictionary<string, object?>
                {
                    ["index"] = c,
                    ["label"] = ColLetter(c),
                };

                // Header value from row 0
                var headerCell = ws.Cells[0, c];
                colInfo["header"] = headerCell?.DisplayStringValue ?? "";

                // Sample data (from row 1 onward)
                var samples = new List<string?>();
                for (int r = 1; r <= Math.Min(sampleRowCount, usedRows - 1); r++)
                {
                    var cell = ws.Cells[r, c];
                    samples.Add(cell?.DisplayStringValue);
                }
                colInfo["sampleValues"] = samples;

                // Detect type from sample
                string colType = "String";
                foreach (var s in samples)
                {
                    if (string.IsNullOrEmpty(s)) continue;
                    if (int.TryParse(s, out _) || decimal.TryParse(s, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out _))
                        colType = "Numeric";
                    else if (DateTime.TryParse(s, System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None, out _))
                        colType = "DateTime";
                    else if (bool.TryParse(s, out _))
                        colType = "Bool";
                }
                colInfo["inferredType"] = colType;

                // Formula pattern detection for this column
                var fPattern = DetectColumnFormulaPattern(ws, c, usedRows);
                if (fPattern != null)
                    colInfo["formulaPattern"] = fPattern;

                columns.Add(colInfo);
            }
            if (columns.Count > 0)
                desc["columns"] = columns;

            // Formula patterns (overall)
            var allPatterns = DetectAllFormulaPatterns(ws, usedRows, usedCols);
            if (allPatterns.Count > 0)
                desc["formulaPatterns"] = allPatterns;

            // Merged cells
            var merged = ws.Cells.MergedCells;
            if (merged != null && merged.Count > 0)
            {
                var mergedList = new List<string>();
                foreach (var m in merged)
                    mergedList.Add($"{CellRefFromIdx(m.StartRow, m.StartColumn)}:{CellRefFromIdx(m.EndRow, m.EndColumn)}");
                desc["mergedCells"] = mergedList;
            }

            // Tables
            if (ws.ListObjects.Count > 0)
            {
                var tables = new List<Dictionary<string, object?>>();
                for (int i = 0; i < ws.ListObjects.Count; i++)
                {
                    var t = ws.ListObjects[i];
                    tables.Add(new Dictionary<string, object?>
                    {
                        ["name"] = t.DisplayName,
                        ["range"] = $"{CellRefFromIdx(t.StartRow, t.StartColumn)}:{CellRefFromIdx(t.EndRow, t.EndColumn)}",
                        ["hasHeaders"] = t.ShowHeaderRow,
                        ["hasTotals"] = t.ShowTotals,
                    });
                }
                desc["tables"] = tables;
            }

            // Charts
            if (ws.Charts.Count > 0)
            {
                var charts = new List<Dictionary<string, object?>>();
                for (int i = 0; i < ws.Charts.Count; i++)
                {
                    var c = ws.Charts[i];
                    charts.Add(new Dictionary<string, object?>
                    {
                        ["name"] = c.Name,
                        ["type"] = c.ChartType.ToString(),
                    });
                }
                desc["charts"] = charts;
            }

            // Comments
            if (ws.Comments.Count > 0)
                desc["commentCount"] = ws.Comments.Count;

            // AutoFilter
            if (!string.IsNullOrEmpty(ws.AutoFilter?.Range))
                desc["autoFilter"] = ws.AutoFilter.Range;

            // Protection
            desc["isProtected"] = ws.Protection?.IsProtected ?? false;

            return desc;
        }

        private Dictionary<string, object?>? DetectColumnFormulaPattern(Worksheet ws, int col, int usedRows)
        {
            // Collect formulas for this column, starting from row 1 (skip header)
            var formulas = new List<(int Row, string Formula)>();
            for (int r = 1; r < usedRows; r++)
            {
                var cell = ws.Cells[r, col];
                if (cell == null) continue;
                var f = cell.Formula;
                if (!string.IsNullOrEmpty(f))
                    formulas.Add((r, f));
            }

            if (formulas.Count == 0) return null;
            if (formulas.Count == 1)
            {
                return new Dictionary<string, object?>
                {
                    ["formula"] = formulas[0].Formula,
                    ["range"] = CellRefFromIdx(formulas[0].Row, col),
                    ["isSingle"] = true,
                };
            }

            // Normalize: replace row numbers with {r} placeholder
            var normalized = formulas.Select(f => (f.Row, Normalized: NormalizeFormulaRow(f.Formula, f.Row))).ToList();

            // Group by normalized form
            var groups = normalized.GroupBy(x => x.Normalized).ToList();

            var results = new List<Dictionary<string, object?>>();
            foreach (var group in groups)
            {
                var rows = group.Select(x => x.Row).OrderBy(r => r).ToList();
                if (rows.Count == 0) continue;
                var range = rows.Count == 1
                    ? CellRefFromIdx(rows[0], col)
                    : $"{CellRefFromIdx(rows[0], col)}:{CellRefFromIdx(rows[^1], col)}";

                results.Add(new Dictionary<string, object?>
                {
                    ["formula"] = group.Key,
                    ["range"] = range,
                    ["count"] = rows.Count,
                    ["example"] = $"{CellRefFromIdx(rows[0], col)}: {formulas.First(f => f.Row == rows[0]).Formula}",
                });
            }

            // Return the dominant pattern
            return results.OrderByDescending(r => (int)r["count"]!).First();
        }

        private List<Dictionary<string, object?>> DetectAllFormulaPatterns(Worksheet ws, int usedRows, int usedCols)
        {
            // Collect all formula cells
            var allFormulas = new List<(int Row, int Col, string Formula)>();
            for (int r = 1; r < usedRows; r++)
            {
                for (int c = 0; c < usedCols; c++)
                {
                    var cell = ws.Cells[r, c];
                    if (cell == null) continue;
                    var f = cell.Formula;
                    if (!string.IsNullOrEmpty(f))
                        allFormulas.Add((r, c, f));
                }
            }

            if (allFormulas.Count == 0) return new List<Dictionary<string, object?>>();

            // Normalize and group
            var normalized = allFormulas.Select(f => (
                f.Row, f.Col,
                Normalized: NormalizeFormulaRow(f.Formula, f.Row)
            )).ToList();

            var groups = normalized.GroupBy(x => x.Normalized).ToList();

            var result = new List<Dictionary<string, object?>>();
            foreach (var group in groups.OrderByDescending(g => g.Count()))
            {
                var cells = group.Select(x => (x.Row, x.Col)).OrderBy(x => x.Row).ThenBy(x => x.Col).ToList();
                if (cells.Count == 0) continue;

                // Find contiguous ranges
                var ranges = new List<string>();
                int startR = cells[0].Row, startC = cells[0].Col;
                int endR = startR, endC = startC;

                for (int i = 1; i < cells.Count; i++)
                {
                    bool adjacent = (cells[i].Row == endR && cells[i].Col == endC + 1) ||
                                    (cells[i].Row == endR + 1 && cells[i].Col == 0);
                    if (adjacent)
                    {
                        endR = cells[i].Row;
                        endC = cells[i].Col;
                    }
                    else
                    {
                        ranges.Add($"{CellRefFromIdx(startR, startC)}:{CellRefFromIdx(endR, endC)}");
                        startR = cells[i].Row; startC = cells[i].Col;
                        endR = startR; endC = startC;
                    }
                }
                ranges.Add($"{CellRefFromIdx(startR, startC)}:{CellRefFromIdx(endR, endC)}");

                var first = allFormulas.First(f => f.Row == cells[0].Row && f.Col == cells[0].Col);

                result.Add(new Dictionary<string, object?>
                {
                    ["formula"] = group.Key,
                    ["range"] = string.Join("; ", ranges.Take(5)),
                    ["count"] = cells.Count,
                    ["example"] = $"{CellRefFromIdx(first.Row, first.Col)}: {first.Formula}",
                });

                if (result.Count >= 10) break; // cap at 10 patterns
            }

            return result;
        }

        private static string NormalizeFormulaRow(string formula, int currentRow)
        {
            // Use regex to find cell references like A1, BC12, $A$1 and normalize row numbers
            return System.Text.RegularExpressions.Regex.Replace(formula,
                @"([A-Za-z]{1,3})(\d+)",
                m => m.Groups[1].Value + "{r}");
        }

        private static string CellRefFromIdx(int row, int col)
        {
            if (row < 0 || col < 0) return "";
            return $"{ColLetter(col)}{row + 1}";
        }

        private static string ColLetter(int col)
        {
            var sb = new System.Text.StringBuilder();
            while (col >= 0)
            {
                sb.Insert(0, (char)('A' + (col % 26)));
                col = col / 26 - 1;
            }
            return sb.ToString();
        }

        // ──────────────────────────────────────────────
        //  Deterministic auto-format (save-time pass)
        // ──────────────────────────────────────────────

        /// <summary>Pastel palette for table titles, rotated per detected table. Light enough to
        /// keep the default dark font readable.</summary>
        private static readonly Color[] TitlePalette =
        {
            Color.FromArgb(255, 221, 235, 247),   // light blue
            Color.FromArgb(255, 226, 239, 218),   // light green
            Color.FromArgb(255, 252, 228, 214),   // light orange
            Color.FromArgb(255, 228, 223, 236),   // light purple
            Color.FromArgb(255, 255, 242, 204),   // light yellow
            Color.FromArgb(255, 217, 226, 243),   // light steel
        };

        /// <summary>Default column width (chars) used when no width is set; a column at this
        /// width (or unset) is auto-fitted at save. Explicitly user-set widths are preserved.</summary>
        private const double DefaultColumnWidth = 8.43;

        /// <summary>
        /// Deterministic, agent-invisible formatting applied at every persist:
        ///   • columns that still carry the DEFAULT width are auto-fitted to their content
        ///     (format-aware: numbers with a number format and formula cells get a safe width,
        ///     so formatted values never render as "###"); columns explicitly adjusted by the
        ///     user/agent are left untouched;
        ///   • text cells forming a contiguous vertical run directly above a number/formula
        ///     cell are styled as table titles (bold + light pastel background, one palette
        ///     color per table) — "general title → column header → data" is all styled, in
        ///     any write order. Cells already styled by the user/agent are never overridden.
        /// Idempotent: styled/fitted once, skipped afterwards.
        /// </summary>
        private void ApplyDeterministicAutoFormat()
        {
            if (_workbook == null) return;
            _bestFitCols.Clear();
            foreach (var ws in _workbook.Worksheets)
                ApplyDeterministicAutoFormatToSheet(ws, _touchedCols, _bestFitCols);
        }

        private static void ApplyDeterministicAutoFormatToSheet(Worksheet ws,
            Dictionary<string, HashSet<int>> touchedCols, Dictionary<string, HashSet<int>> bestFitCols)
        {
            int lastRow = FindLastUsedRow(ws);
            if (lastRow < 0) return;

            // Phase 1 — per column: content scan + title runs.
            // runTop/runBottom: for every cell that is part of a title run, the run's first
            // and last row (bottom = the row directly above a data cell).
            var runTop = new Dictionary<(int Row, int Col), int>();
            var runBottom = new Dictionary<(int Row, int Col), int>();

            for (int c = 0; c < 500; c++)
            {
                bool hasContent = false;
                var textSet = new HashSet<int>();
                var dataRows = new List<int>();
                for (int r = 0; r <= lastRow; r++)
                {
                    var cell = ws.Cells[r, c];
                    if (cell == null || cell.Type == CellValueType.IsNull) continue;
                    hasContent = true;
                    if (!string.IsNullOrEmpty(cell.Formula) || cell.Type is CellValueType.IsNumeric
                        or CellValueType.IsDateTime or CellValueType.IsBool)
                        dataRows.Add(r);
                    else if (cell.Type == CellValueType.IsString && !string.IsNullOrEmpty(cell.DisplayStringValue))
                        textSet.Add(r);
                }
                if (!hasContent) continue;

                // bestFit for columns the agent TOUCHED this session and that still carry the
                // DEFAULT width: a width the user set is untouchable (the user is always right).
                // The flag tells the opening application to auto-fit the column to its content
                // (formulas included — the app computes the results, we cannot).
                if (touchedCols.TryGetValue(ws.Name, out var touched) && touched.Contains(c))
                {
                    var width = ws.Cells.Columns[c].Width;
                    if (width == null || Math.Abs(width.Value - DefaultColumnWidth) < 0.01)
                        AddBestFitColumn(bestFitCols, ws.Name, c);
                }

                // Title runs: for each data cell, the contiguous text cells immediately above.
                foreach (var dr in dataRows)
                {
                    int top = dr - 1;
                    while (top >= 0 && textSet.Contains(top)) top--;
                    int bottom = dr - 1;
                    if (top >= bottom) continue;   // no text run above this data cell
                    for (int r = top + 1; r <= bottom; r++)
                    {
                        runTop[(r, c)] = top + 1;
                        runBottom[(r, c)] = bottom;
                    }
                }
            }

            if (runBottom.Count == 0) return;

            // Phase 2 — table grouping: a table's title row = maximal horizontal run of cells
            // that are the BOTTOM of their vertical run (data directly below). Each table gets
            // the next palette color; the cells above in the run inherit it.
            var titleColor = new Dictionary<(int Row, int Col), Color>();
            int tableIdx = 0;
            for (int r = 0; r <= lastRow; r++)
            {
                int c = 0;
                while (c < 500)
                {
                    bool isBottom = runBottom.TryGetValue((r, c), out var b) && b == r;
                    if (!isBottom) { c++; continue; }
                    int start = c;
                    while (c < 500 && runBottom.TryGetValue((r, c), out var b2) && b2 == r) c++;
                    var color = TitlePalette[tableIdx++ % TitlePalette.Length];
                    for (int cc = start; cc < c; cc++)
                    {
                        var top = runTop[(r, cc)];
                        for (int rr = top; rr <= r; rr++)
                            titleColor[(rr, cc)] = color;
                    }
                }
            }

            // Phase 3 — apply the style, skipping cells the user/agent already styled.
            foreach (var ((r, c), color) in titleColor)
            {
                var cell = ws.Cells[r, c];
                if (cell == null) continue;
                var style = cell.GetStyle();
                if (style.Pattern != FillPattern.None || style.Font.IsBold) continue;
                style.Font.IsBold = true;
                style.Pattern = FillPattern.Solid;
                style.ForegroundColor = color;
                cell.SetStyle(style);
            }
        }

        /// <summary>Records that the agent wrote to or styled a column this session — the marker
        /// the auto-format pass uses to decide which columns get bestFit.</summary>
        private void MarkColumnTouched(string sheetName, int col)
        {
            if (!_touchedCols.TryGetValue(sheetName, out var set))
                _touchedCols[sheetName] = set = new HashSet<int>();
            set.Add(col);
        }

        /// <summary>Adds a column to the bestFit set of a worksheet (populated at save time).</summary>
        private static void AddBestFitColumn(Dictionary<string, HashSet<int>> bestFitCols, string sheetName, int col)
        {
            if (!bestFitCols.TryGetValue(sheetName, out var set))
                bestFitCols[sheetName] = set = new HashSet<int>();
            set.Add(col);
        }

        /// <summary>Injects the OOXML bestFit flag into the saved workbook: for every column the
        /// agent touched with a default width, the <c>&lt;col&gt;</c> element is written WITHOUT a
        /// width and with bestFit="1", so the OPENING APPLICATION auto-fits the column to its
        /// rendered content (formula results included — the app computes them, we cannot). Columns
        /// with a user-set width keep it untouched. Runs after save, on the temp package, exactly
        /// like the chart-cache patch; the file is committed only if every part still parses.</summary>
        private void InjectBestFit(string filePath)
        {
            if (_bestFitCols.Count == 0 || _workbook == null || !File.Exists(filePath)) return;
            try
            {
                // Worksheet index → touched default-width columns (fork writes sheets in order).
                var colsBySheet = new Dictionary<int, HashSet<int>>();
                for (int i = 0; i < _workbook.Worksheets.Count; i++)
                    if (_bestFitCols.TryGetValue(_workbook.Worksheets[i].Name, out var cols) && cols.Count > 0)
                        colsBySheet[i] = cols;
                if (colsBySheet.Count == 0) return;

                // Same all-in-memory rewrite pattern as PatchChartCaches: read every entry,
                // patch the targeted worksheet parts, write a fresh package and swap.
                var entries = new List<(string Name, byte[] Data)>();
                using (var zip = ZipFile.OpenRead(filePath))
                    foreach (var e in zip.Entries)
                    {
                        using var ms = new MemoryStream();
                        using (var s = e.Open()) s.CopyTo(ms);
                        entries.Add((e.FullName, ms.ToArray()));
                    }

                var changed = false;
                for (int i = 0; i < entries.Count; i++)
                {
                    var (name, data) = entries[i];
                    var m = System.Text.RegularExpressions.Regex.Match(name, @"^xl/worksheets/sheet(\d+)\.xml$");
                    if (!m.Success) continue;
                    if (!colsBySheet.TryGetValue(int.Parse(m.Groups[1].Value) - 1, out var cols)) continue;
                    var text = System.Text.Encoding.UTF8.GetString(data);
                    var patched = InjectBestFitIntoSheetXml(text, cols);
                    if (patched == text) continue;
                    entries[i] = (name, System.Text.Encoding.UTF8.GetBytes(patched));
                    changed = true;
                }
                if (!changed) return;

                var tmp = filePath + ".tmp";
                using (var outZip = ZipFile.Open(tmp, ZipArchiveMode.Create))
                    foreach (var (name, data) in entries)
                    {
                        var entry = outZip.CreateEntry(name);
                        using var s = entry.Open();
                        s.Write(data, 0, data.Length);
                    }
                File.Delete(filePath);
                File.Move(tmp, filePath);
            }
            catch (Exception ex)
            {
                Log.LogStep($"SpreadsheetTool.InjectBestFit: failed — {ex.Message}");
            }
        }

        /// <summary>Rewrites the &lt;cols&gt; section of a worksheet part: keeps the existing (user
        /// width) columns and adds a bestFit-only &lt;col&gt; for each touched default-width column,
        /// sorted by min. Creates &lt;cols&gt; when absent (before &lt;sheetData&gt;).</summary>
        private static string InjectBestFitIntoSheetXml(string xml, HashSet<int> cols)
        {
            var colsMatch = System.Text.RegularExpressions.Regex.Match(xml, "<cols>.*?</cols>");
            var existing = new List<string>();
            if (colsMatch.Success)
            {
                foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(colsMatch.Value, "<col [^>]*/>"))
                    existing.Add(m.Value);
            }

            var bestFit = cols.Select(c => $"<col min=\"{c + 1}\" max=\"{c + 1}\" bestFit=\"1\" />").ToList();
            var all = existing.Concat(bestFit)
                .Select(colXml =>
                {
                    var min = System.Text.RegularExpressions.Regex.Match(colXml, "min=\"(\\d+)\"").Groups[1].Value;
                    return (Min: int.Parse(min), Xml: colXml);
                })
                .OrderBy(x => x.Min)
                .Select(x => x.Xml)
                .ToList();
            var newCols = "<cols>" + string.Concat(all) + "</cols>";

            if (colsMatch.Success)
            {
                var result = xml.Replace(colsMatch.Value, newCols);
                return result == xml ? xml : result;
            }
            // No <cols> section: insert it right before <sheetData> (schema order).
            var marker = "<sheetData>";
            var idx = xml.IndexOf(marker, StringComparison.Ordinal);
            return idx < 0 ? xml : xml.Insert(idx, newCols);
        }

        /// <summary>Writes the in-memory workbook to <paramref name="targetPath"/> through a temp
        /// file, patches the chart caches, and commits the file ONLY when every XML part parses
        /// (the same bar Excel/LibreOffice apply on open). The target file is never replaced by
        /// an unvalidated package. Returns null on success, else the validation/save error.</summary>
        private string? PersistValidated(string targetPath)
        {
            var tmp = targetPath + ".validating";
            try
            {
                _workbook!.Save(tmp);
                PatchChartCaches(tmp);
                InjectBestFit(tmp);
                var validationError = ValidateXmlParts(tmp);
                if (validationError != null)
                    return validationError;
                File.Copy(tmp, targetPath, overwrite: true);
                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
            }
        }

        /// <summary>Strict well-formedness check of every XML part in a saved workbook package.
        /// The serializer writes valid XML by construction, but the chart cache patch rewrites
        /// parts — this gate guarantees the file on disk always opens in Excel/LibreOffice.</summary>
        private static string? ValidateXmlParts(string filePath)
        {
            try
            {
                using var zip = ZipFile.OpenRead(filePath);
                foreach (var e in zip.Entries)
                {
                    if (!e.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) continue;
                    using var s = e.Open();
                    var doc = new System.Xml.XmlDocument();
                    try { doc.Load(s); }
                    catch (System.Xml.XmlException ex)
                    {
                        return $"invalid XML in '{e.FullName}': {ex.Message}";
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                return $"cannot validate '{filePath}': {ex.Message}";
            }
        }

        /// <summary>Populates the chart numCache entries of a saved workbook: the FOSS chart
        /// template always writes an EMPTY cache (ptCount=0), which renders as a blank chart in
        /// LibreOffice/Excel. The values are read from the referenced cells and written into the
        /// saved package — deterministic, no app-side refresh needed.</summary>
        private void PatchChartCaches(string filePath)
        {
            if (_workbook == null || !File.Exists(filePath)) return;
            try
            {
                var entries = new List<(string Name, byte[] Data)>();
                using (var zip = ZipFile.OpenRead(filePath))
                    foreach (var e in zip.Entries)
                    {
                        using var ms = new MemoryStream();
                        using (var s = e.Open()) s.CopyTo(ms);
                        entries.Add((e.FullName, ms.ToArray()));
                    }

                var changed = false;
                for (int i = 0; i < entries.Count; i++)
                {
                    var (name, data) = entries[i];
                    if (!name.StartsWith("xl/charts/") || !name.EndsWith(".xml")) continue;
                    var text = System.Text.Encoding.UTF8.GetString(data);
                    var patched = PopulateChartCacheXml(text);
                    if (patched != text)
                    {
                        entries[i] = (name, System.Text.Encoding.UTF8.GetBytes(patched));
                        changed = true;
                    }
                }
                if (!changed) return;

                var tmp = filePath + ".tmp";
                using (var outZip = ZipFile.Open(tmp, ZipArchiveMode.Create))
                    foreach (var (name, data) in entries)
                    {
                        var entry = outZip.CreateEntry(name);
                        using var s = entry.Open();
                        s.Write(data, 0, data.Length);
                    }
                File.Delete(filePath);
                File.Move(tmp, filePath);
            }
            catch (Exception ex)
            {
                Log.LogStep($"SpreadsheetTool.PatchChartCaches: failed — {ex.Message}");
            }
        }

        // The closing </c:numCache> is part of the match so the whole element is replaced
        // atomically: a replacement ending at <c:ptCount val="0"/> left the original
        // </c:numCache></c:numRef> tail behind, producing a duplicate closing tag and a
        // malformed chart XML that Excel/LibreOffice refuse to render.
        private static readonly System.Text.RegularExpressions.Regex ChartCachePattern = new(
            @"<c:f>(?<range>[^<]+)</c:f>\s*<c:numCache><c:formatCode>[^<]*</c:formatCode><c:ptCount val=""0""/></c:numCache>",
            System.Text.RegularExpressions.RegexOptions.Compiled);

        /// <summary>The fork's single-series template, verbatim: used to detect charts we created
        /// (one &lt;c:ser&gt; with only a val and an empty numCache) so preserved charts loaded
        /// from existing files are never restructured.</summary>
        private static readonly System.Text.RegularExpressions.Regex ChartSeriesPattern = new(
            @"<c:ser><c:idx val=""0""/><c:order val=""0""/><c:val><c:numRef><c:f>(?<range>[^<]+)</c:f><c:numCache><c:formatCode>[^<]*</c:formatCode><c:ptCount val=""0""/></c:numCache></c:numRef></c:val></c:ser>",
            System.Text.RegularExpressions.RegexOptions.Compiled);

        /// <summary>Chart families whose series carry &lt;c:cat&gt; + &lt;c:val&gt; (bar/line/area/
        /// pie/doughnut/radar, incl. 3D). Scatter/bubble/stock/surface use different series
        /// layouts and only get their caches filled.</summary>
        private static readonly System.Text.RegularExpressions.Regex ChartFamilyPattern = new(
            @"<c:(bar|line|area|pie|doughnut|radar)[A-Za-z0-9]*Chart>",
            System.Text.RegularExpressions.RegexOptions.Compiled);

        /// <summary>Populates the empty caches of every chart in a saved workbook. Charts we
        /// created from the fork template get restructured: a 2D data block becomes one series
        /// per column with the first row as series names and the first column as categories
        /// (a 1D range stays a single series); all numCache/strCache values are filled from the
        /// referenced cells — deterministic, no app-side refresh needed.</summary>
        private string PopulateChartCacheXml(string xml)
        {
            if (ChartFamilyPattern.IsMatch(xml))
            {
                try
                {
                    var restructured = ChartSeriesPattern.Replace(xml, m =>
                    {
                        var series = BuildSeriesXml(m.Groups["range"].Value);
                        return series ?? m.Value;
                    });
                    if (restructured != xml) return restructured;
                }
                catch
                {
                    // fall through to the cache-only fill below
                }
            }

            // Cache-only path: the remaining families (scatter/bubble/stock/…) and preserved
            // charts whose empty cache was not restructured above still need their values.
            return ChartCachePattern.Replace(xml, m =>
            {
                var pts = BuildChartPointsXml(m.Groups["range"].Value);
                return pts != null
                    ? $"<c:f>{m.Groups["range"].Value}</c:f><c:numCache><c:formatCode>General</c:formatCode>{pts}</c:numCache>"
                    : m.Value;
            });
        }

        /// <summary>Builds the &lt;c:ser&gt; block(s) for a chart range. 2D block → categories
        /// (first column, below the header row) + one series per remaining column with the name
        /// from the header row; 1D range → a single series. Returns null when no numeric value
        /// can be read (chart stays as-is).</summary>
        private string? BuildSeriesXml(string range)
        {
            var bang = range.LastIndexOf('!');
            var sheetName = bang >= 0 ? range[..bang].Trim().Trim('\'') : null;
            var cells = (bang >= 0 ? range[(bang + 1)..] : range).Replace("$", "");
            var ws = sheetName == null ? _workbook!.Worksheets[0] : FindSheet(sheetName);
            if (ws == null) return null;
            var p = TryParseRange(cells);
            if (!p.Ok) return null;
            var r1 = p.R1; var c1 = p.C1;
            var r2 = Math.Min(p.R2, r1 + 1000);
            var c2 = Math.Min(p.C2, c1 + 100);
            var sheetRef = "'" + (sheetName ?? "").Replace("'", "''") + "'";

            bool isMatrix = (r2 - r1 + 1) >= 2 && (c2 - c1 + 1) >= 2;
            int serCount = isMatrix ? c2 - c1 : 1;   // first column becomes the categories

            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < serCount; i++)
            {
                int sCol = isMatrix ? c1 + 1 + i : c1;
                sb.Append("<c:ser><c:idx val=\"").Append(i).Append("\"/><c:order val=\"").Append(i).Append("\"/>");

                if (isMatrix)
                {
                    // Series name from the header row of the series column.
                    var name = ws.Cells[r1, sCol]?.StringValue ?? "";
                    var nameRef = $"{sheetRef}!${ColLetter(sCol)}${r1 + 1}";
                    sb.Append("<c:tx><c:strRef><c:f>").Append(nameRef).Append("</c:f>")
                      .Append("<c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>")
                      .Append(EscapeXml(name)).Append("</c:v></c:pt></c:strCache></c:strRef></c:tx>");

                    // Categories from the first column, rows below the header row.
                    var catPts = BuildTextPointsXml(ws, r1 + 1, c1, r2, c1);
                    if (catPts == null) return null;
                    var catRef = $"{sheetRef}!${ColLetter(c1)}${r1 + 2}:${ColLetter(c1)}${r2 + 1}";
                    sb.Append("<c:cat><c:strRef><c:f>").Append(catRef).Append("</c:f><c:strCache>")
                      .Append(catPts).Append("</c:strCache></c:strRef></c:cat>");
                }

                var numPts = isMatrix
                    ? BuildNumberPointsXml(ws, r1 + 1, sCol, r2, sCol)
                    : BuildNumberPointsXml(ws, r1, c1, r2, c2);
                if (numPts == null) return null;
                var valRef = isMatrix
                    ? $"{sheetRef}!${ColLetter(sCol)}${r1 + 2}:${ColLetter(sCol)}${r2 + 1}"
                    : $"{sheetRef}!${ColLetter(c1)}${r1 + 1}:${ColLetter(c2)}${r2 + 1}";
                sb.Append("<c:val><c:numRef><c:f>").Append(valRef).Append("</c:f>")
                  .Append("<c:numCache><c:formatCode>General</c:formatCode>").Append(numPts)
                  .Append("</c:numCache></c:numRef></c:val></c:ser>");
            }
            return sb.ToString();
        }

        /// <summary>Renders the &lt;c:ptCount&gt; + &lt;c:pt&gt; rows of a strCache for a cell range.</summary>
        private static string? BuildTextPointsXml(Worksheet ws, int r1, int c1, int r2, int c2)
        {
            var sb = new System.Text.StringBuilder();
            int n = 0;
            for (int r = r1; r <= r2; r++)
                for (int c = c1; c <= c2; c++)
                {
                    var s = ws.Cells[r, c]?.StringValue ?? "";
                    sb.Append("<c:pt idx=\"").Append(n).Append("\"><c:v>").Append(EscapeXml(s)).Append("</c:v></c:pt>");
                    n++;
                }
            if (n == 0) return null;
            return $"<c:ptCount val=\"{n}\"/>{sb}";
        }

        /// <summary>Renders the &lt;c:ptCount&gt; + &lt;c:pt&gt; rows of a numCache for a cell range,
        /// keeping only cells that hold numbers. Values are read from the boxed <c>Cell.Value</c>
        /// (never from display strings, which are style/culture dependent). Formula cells without
        /// a cached value are skipped.</summary>
        private static string? BuildNumberPointsXml(Worksheet ws, int r1, int c1, int r2, int c2)
        {
            var sb = new System.Text.StringBuilder();
            int n = 0;
            for (int r = r1; r <= r2; r++)
                for (int c = c1; c <= c2; c++)
                {
                    if (!TryFormatNumber(ws.Cells[r, c]?.Value, out var num)) continue;
                    sb.Append("<c:pt idx=\"").Append(n).Append("\"><c:v>")
                      .Append(num).Append("</c:v></c:pt>");
                    n++;
                }
            if (n == 0) return null;
            return $"<c:ptCount val=\"{n}\"/>{sb}";
        }

        /// <summary>Formats a boxed cell value as an invariant number string, or returns false
        /// for non-numeric values (text, booleans, DateTime, formula cells without a cached
        /// value).</summary>
        private static bool TryFormatNumber(object? value, out string number)
        {
            switch (value)
            {
                case byte b: number = b.ToString(CultureInfo.InvariantCulture); return true;
                case short s: number = s.ToString(CultureInfo.InvariantCulture); return true;
                case int i: number = i.ToString(CultureInfo.InvariantCulture); return true;
                case long l: number = l.ToString(CultureInfo.InvariantCulture); return true;
                case float f: number = f.ToString(CultureInfo.InvariantCulture); return true;
                case double d: number = d.ToString(CultureInfo.InvariantCulture); return true;
                case decimal m: number = m.ToString(CultureInfo.InvariantCulture); return true;
                default: number = string.Empty; return false;
            }
        }

        /// <summary>Reads the numeric values of a chart range (sheet!A1:B5) and renders the
        /// numCache points XML. Text/formula cells without a cached value are skipped.</summary>
        private string? BuildChartPointsXml(string range)
        {
            try
            {
                var bang = range.LastIndexOf('!');
                var sheetName = bang >= 0 ? range[..bang].Trim().Trim('\'') : null;
                var cells = (bang >= 0 ? range[(bang + 1)..] : range).Replace("$", "");
                var ws = sheetName == null ? _workbook!.Worksheets[0] : FindSheet(sheetName);
                if (ws == null) return null;
                var p = TryParseRange(cells);
                if (!p.Ok) return null;
                var r1 = p.R1; var c1 = p.C1;
                var r2 = Math.Min(p.R2, r1 + 1000);
                var c2 = Math.Min(p.C2, c1 + 100);

                var vals = new List<string>();
                for (int r = r1; r <= r2; r++)
                    for (int c = c1; c <= c2; c++)
                    {
                        if (!TryFormatNumber(ws.Cells[r, c]?.Value, out var num)) continue;
                        vals.Add(num);
                    }
                if (vals.Count == 0) return null;

                var sb = new System.Text.StringBuilder($"<c:ptCount val=\"{vals.Count}\"/>");
                for (int i = 0; i < vals.Count; i++)
                    sb.Append("<c:pt idx=\"").Append(i).Append("\"><c:v>").Append(vals[i]).Append("</c:v></c:pt>");
                return sb.ToString();
            }
            catch
            {
                return null;
            }
        }

        private Worksheet? FindSheet(string name)
        {
            for (int i = 0; i < _workbook.Worksheets.Count; i++)
            {
                if (_workbook.Worksheets[i].Name == name)
                    return _workbook.Worksheets[i];
            }
            return null;
        }

        /// <summary>Guards the "A1" or "A1:C5" cellOrRange parameter shared by the style methods.
        /// Returns null when valid, else the error message.</summary>
        private static string? ValidateCellOrRange(string cellOrRange)
        {
            if (cellOrRange.Contains(':'))
                return TryParseRange(cellOrRange).Ok ? null : TryParseRange(cellOrRange).Error;
            var p = TryParseCellRef(cellOrRange);
            return p.Ok ? null : p.Error;
        }

        /// <summary>Strict A1-reference parser used as a guard at every public entry point.
        /// Returns an error message instead of throwing, so the agent gets certain feedback
        /// on the exact parameter that was rejected.</summary>
        private static (bool Ok, int Row, int Col, string Error) TryParseCellRef(string refStr)
        {
            if (string.IsNullOrEmpty(refStr)) return (false, 0, 0, $"invalid cell reference '{refStr}'");
            int col = 0, i = 0;
            while (i < refStr.Length && char.IsLetter(refStr[i]))
            {
                col = col * 26 + (char.ToUpper(refStr[i]) - 'A' + 1);
                i++;
            }
            if (i == 0 || i >= refStr.Length || !int.TryParse(refStr.AsSpan(i), out var row)
                || row < 1 || row > 1_048_576 || col < 1 || col > 16_384)
                return (false, 0, 0, $"invalid cell reference '{refStr}'");
            return (true, row - 1, col - 1, "");
        }

        /// <summary>Strict "A1:C5" range parser (single cell allowed). Validates order and caps the
        /// area so a bad agent parameter can never trigger an unbounded operation.</summary>
        private static (bool Ok, int R1, int C1, int R2, int C2, string Error) TryParseRange(string range)
        {
            if (string.IsNullOrEmpty(range)) return (false, 0, 0, 0, 0, $"invalid range '{range}'");
            var parts = range.Split(':');
            if (parts.Length > 2) return (false, 0, 0, 0, 0, $"invalid range '{range}'");
            var a = TryParseCellRef(parts[0]);
            if (!a.Ok) return (false, 0, 0, 0, 0, a.Error);
            var b = TryParseCellRef(parts.Length == 2 ? parts[1] : parts[0]);
            if (!b.Ok) return (false, 0, 0, 0, 0, b.Error);
            var r1 = Math.Min(a.Row, b.Row); var c1 = Math.Min(a.Col, b.Col);
            var r2 = Math.Max(a.Row, b.Row); var c2 = Math.Max(a.Col, b.Col);
            long area = (long)(r2 - r1 + 1) * (c2 - c1 + 1);
            if (area > MaxCellArea) return (false, 0, 0, 0, 0, $"range '{range}' too large ({area} cells, max {MaxCellArea})");
            return (true, r1, c1, r2, c2, "");
        }

        /// <summary>Escapes a value for insertion into chart XML (&lt;c:v&gt; text).</summary>
        private static string EscapeXml(string s) =>
            s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");

        /// <summary>0-based index of the last row carrying a value OR a formula (formula-only
        /// cells display as empty without recalculation, so a display-only scan would stop early
        /// and a later append would overwrite formula rows). Scans all columns, capped like
        /// ScanNonEmptyColumns. Returns -1 for an empty sheet.</summary>
        private static int FindLastUsedRow(Worksheet ws)
        {
            int lastUsedRow = -1;
            for (int r = 0; r < 5000; r++)
            {
                bool any = false;
                for (int c = 0; c < 500; c++)
                {
                    var cell = ws.Cells[r, c];
                    if (cell != null && (!string.IsNullOrEmpty(cell.DisplayStringValue) || !string.IsNullOrEmpty(cell.Formula)))
                    { any = true; break; }
                }
                if (any) lastUsedRow = r;
                else if (r > lastUsedRow + 5) break; // 5 empty rows → stop
            }
            return lastUsedRow;
        }

        private static void SetCellValueAuto(Cell cell, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                cell.PutValue(string.Empty);
                return;
            }

            // A string starting with "=" is a formula: store it as such. PutValue would keep
            // it as literal text, which spreadsheet apps render with a leading apostrophe
            // (the "text that looks like a formula" convention) — a broken formula.
            if (value.StartsWith('='))
            {
                cell.Formula = value;
                return;
            }

            if (bool.TryParse(value, out var bVal)) { cell.PutValue(bVal); return; }
            // Numbers and dates must not silently become text when written in the user's locale
            // ("1,5" decimal comma, "15/01/2026" dd/MM/yyyy): text cells are excluded from
            // formulas and from chart data, so an Italian-locale value would corrupt both.
            if (TryParseNumber(value, out var dVal)) { cell.PutValue(dVal); return; }
            if (TryParseDate(value, out var dtVal)) { cell.PutValue(dtVal); return; }

            cell.PutValue(value);
        }

        /// <summary>Deterministic decimal parsing: the LAST separator decides the culture.
        /// Comma last → European decimal comma ("1,5", "1.234,56" → it-IT); dot last →
        /// invariant ("1.5", "1,234.56"). Never ambiguous, both locale conventions work.</summary>
        private static bool TryParseNumber(string value, out decimal result)
        {
            var styles = System.Globalization.NumberStyles.Number;
            var lastDot = value.LastIndexOf('.');
            var lastComma = value.LastIndexOf(',');
            var culture = lastComma > lastDot
                ? System.Globalization.CultureInfo.GetCultureInfo("it-IT")
                : System.Globalization.CultureInfo.InvariantCulture;
            return decimal.TryParse(value, styles, culture, out result);
        }

        private static bool TryParseDate(string value, out DateTime result)
        {
            if (DateTime.TryParse(value, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out result)) return true;
            if (DateTime.TryParse(value, System.Globalization.CultureInfo.GetCultureInfo("it-IT"),
                System.Globalization.DateTimeStyles.None, out result)) return true;
            return false;
        }

        private static Color ParseColor(string hex)
        {
            if (hex.StartsWith("#")) hex = hex.Substring(1);
            if (hex.Length == 6)
            {
                var r = Convert.ToByte(hex.Substring(0, 2), 16);
                var g = Convert.ToByte(hex.Substring(2, 2), 16);
                var b = Convert.ToByte(hex.Substring(4, 2), 16);
                return Color.FromArgb(255, r, g, b);
            }
            return Color.FromArgb(255, 0, 0, 0);
        }

        private static Aspose.Cells_FOSS.HorizontalAlignmentType ParseHorizontalAlignment(string value)
        {
            return value.ToLowerInvariant() switch
            {
                "general" => Aspose.Cells_FOSS.HorizontalAlignmentType.General,
                "left" => Aspose.Cells_FOSS.HorizontalAlignmentType.Left,
                "center" => Aspose.Cells_FOSS.HorizontalAlignmentType.Center,
                "right" => Aspose.Cells_FOSS.HorizontalAlignmentType.Right,
                "fill" => Aspose.Cells_FOSS.HorizontalAlignmentType.Fill,
                "justify" => Aspose.Cells_FOSS.HorizontalAlignmentType.Justify,
                "centeracrossselection" => Aspose.Cells_FOSS.HorizontalAlignmentType.CenterContinuous,
                "centercontinuous" => Aspose.Cells_FOSS.HorizontalAlignmentType.CenterContinuous,
                "distributed" => Aspose.Cells_FOSS.HorizontalAlignmentType.Distributed,
                _ => Aspose.Cells_FOSS.HorizontalAlignmentType.General,
            };
        }

        private static VerticalAlignmentType ParseVerticalAlignment(string value)
        {
            return value.ToLowerInvariant() switch
            {
                "top" => VerticalAlignmentType.Top,
                "center" => VerticalAlignmentType.Center,
                "bottom" => VerticalAlignmentType.Bottom,
                "justify" => VerticalAlignmentType.Justify,
                "distributed" => VerticalAlignmentType.Distributed,
                _ => VerticalAlignmentType.Bottom,
            };
        }

        /// <summary>Normalizes an agent-supplied number format. Models often wrap the format in
        /// quotes ("$#,##0.00") — a quoted string in Excel format syntax is a LITERAL, so those
        /// quotes would make every cell display the format text instead of the value. The quotes
        /// are stripped ONLY when the inner content looks like a format code (contains #, 0 or ?
        /// placeholders); a legit quoted-literal format (e.g. "Total: "0.00) is preserved.</summary>
        private static string NormalizeNumberFormat(string format)
        {
            var f = format.Trim();
            if (f.Length >= 2 && f[0] == '"' && f[^1] == '"')
            {
                var inner = f[1..^1];
                if (inner.Any(ch => ch is '#' or '0' or '?'))
                    return inner;
            }
            return f;
        }

        private static BorderStyleType ParseBorderStyle(string value)
        {
            return value.ToLowerInvariant() switch
            {
                "none" => BorderStyleType.None,
                "thin" => BorderStyleType.Thin,
                "medium" => BorderStyleType.Medium,
                "thick" => BorderStyleType.Thick,
                "dotted" => BorderStyleType.Dotted,
                "dashed" => BorderStyleType.Dashed,
                "double" => BorderStyleType.Double,
                "hair" => BorderStyleType.Hair,
                _ => BorderStyleType.Thin,
            };
        }

        private static string[] ParseBorderSides(string side)
        {
            return side.ToLowerInvariant() switch
            {
                "all" => new[] { "Top", "Bottom", "Left", "Right" },
                "outline" => new[] { "Top", "Bottom", "Left", "Right" },
                "inside" => new[] { "Top", "Bottom", "Left", "Right" },
                "top" => new[] { "Top" },
                "bottom" => new[] { "Bottom" },
                "left" => new[] { "Left" },
                "right" => new[] { "Right" },
                _ => new[] { "Top", "Bottom", "Left", "Right" },
            };
        }

        private static ChartType? ParseChartType(string value)
        {
            return value.ToLowerInvariant() switch
            {
                "bar" => ChartType.Bar,
                "column" => ChartType.Column,
                "line" => ChartType.Line,
                "area" => ChartType.Area,
                "pie" => ChartType.Pie,
                "doughnut" => ChartType.Doughnut,
                "scatter" => ChartType.Scatter,
                "bubble" => ChartType.Bubble,
                "radar" => ChartType.Radar,
                "stock" => ChartType.Stock,
                "bar3d" => ChartType.Bar3D,
                "column3d" => ChartType.Column3D,
                "line3d" => ChartType.Line3D,
                "area3d" => ChartType.Area3D,
                "pie3d" => ChartType.Pie3D,
                "waterfall" => ChartType.Waterfall,
                "treemap" => ChartType.Treemap,
                "sunburst" => ChartType.Sunburst,
                "histogram" => ChartType.Histogram,
                "funnel" => ChartType.Funnel,
                "map" => ChartType.Map,
                _ => null,
            };
        }

        private static PaperSizeType ParsePaperSize(string value)
        {
            return value.ToLowerInvariant() switch
            {
                "a4" => PaperSizeType.PaperA4,
                "a3" => PaperSizeType.PaperA3,
                "a5" => PaperSizeType.PaperA5,
                "letter" => PaperSizeType.PaperLetter,
                "legal" => PaperSizeType.PaperLegal,
                "tabloid" => PaperSizeType.PaperTabloid,
                "executive" => PaperSizeType.PaperExecutive,
                "statement" => PaperSizeType.PaperStatement,
                "b4" => PaperSizeType.PaperB4,
                "b5" => PaperSizeType.PaperB5,
                _ => PaperSizeType.PaperA4,
            };
        }

        private static FormatConditionType ParseFormatConditionType(string value)
        {
            return value.ToLowerInvariant() switch
            {
                "cellvalue" => FormatConditionType.CellValue,
                "expression" or "formula" => FormatConditionType.Expression,
                "containstext" => FormatConditionType.ContainsText,
                "notcontainstext" => FormatConditionType.NotContainsText,
                "beginswith" => FormatConditionType.BeginsWith,
                "endswith" => FormatConditionType.EndsWith,
                "timeperiod" => FormatConditionType.TimePeriod,
                "duplicatevalues" or "duplicate" => FormatConditionType.DuplicateValues,
                "uniquevalues" or "unique" => FormatConditionType.UniqueValues,
                "top10" => FormatConditionType.Top10,
                "bottom10" => FormatConditionType.Bottom10,
                "aboveaverage" => FormatConditionType.AboveAverage,
                "belowaverage" => FormatConditionType.BelowAverage,
                "colorscale" => FormatConditionType.ColorScale,
                "databar" => FormatConditionType.DataBar,
                "iconset" => FormatConditionType.IconSet,
                _ => FormatConditionType.CellValue,
            };
        }

        private static OperatorType ParseOperatorType(string value)
        {
            return value.ToLowerInvariant() switch
            {
                "between" => OperatorType.Between,
                "notbetween" => OperatorType.NotBetween,
                "equal" => OperatorType.Equal,
                "notequal" => OperatorType.NotEqual,
                "greaterthan" => OperatorType.GreaterThan,
                "greaterorequal" or "greaterthanorequal" => OperatorType.GreaterOrEqual,
                "lessthan" => OperatorType.LessThan,
                "lessorequal" or "lessthanorequal" => OperatorType.LessOrEqual,
                "none" => OperatorType.None,
                _ => OperatorType.None,
            };
        }
    }
}
