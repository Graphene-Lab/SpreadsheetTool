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

        /// <summary>Columns whose width was already auto-set in this session, per worksheet.
        /// First use of a fresh column gets an auto width; later edits keep it (deterministic
        /// assist — the agent should not have to estimate column widths).</summary>
        private readonly Dictionary<string, HashSet<int>> _autoWidthCols = new();

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
        /// Path to an existing .xlsx file (Unix style, e.g. "/folder/file.xlsx"), relative
        /// to the workspace root (the sandbox). Absolute Windows paths are accepted only
        /// if they descend from the sandbox root.
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
        /// <returns>True if the file was opened successfully.</returns>
        public bool Open(string filePath)
        {
            try
            {
                _workbook?.Dispose();
                _filePath = SandboxPath.Resolve(filePath);
                _workbook = new Workbook(_filePath);
                // Columns that already carry content keep their width on later edits: record
                // them so the first-write auto width does not override an existing layout.
                _autoWidthCols.Clear();
                foreach (var ws in _workbook.Worksheets)
                    _autoWidthCols[ws.Name] = ScanNonEmptyColumns(ws);
                Log.LogStep($"SpreadsheetTool.Open: opened '{_filePath}'");
                return true;
            }
            catch (Exception ex)
            {
                Log.LogStep($"SpreadsheetTool.Open: failed '{filePath}': {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Creates a new XLSX workbook with one default worksheet ("Sheet1") on THIS instance.
        /// Must be an instance method (not a static factory): the agent loop keeps ONE shared
        /// instance in its agents dictionary, so a static Create returning a brand-new agent
        /// discarded the workbook and every later edit failed with a NullReferenceException.
        /// </summary>
        /// <param name="filePath">
        /// Path where the new .xlsx file will be saved (Unix style, e.g. "/folder/file.xlsx"),
        /// relative to the sandbox workspace root.
        /// </param>
        /// <returns>True when the workbook was created and saved.</returns>
        public bool Create(string filePath)
        {
            var resolved = SandboxPath.Resolve(filePath);
            _workbook?.Dispose();
            _workbook = new Workbook();
            _autoWidthCols.Clear();
            _workbook.Save(resolved);
            _filePath = resolved;
            Log.LogStep($"SpreadsheetTool.Create: created '{resolved}'");
            return true;
        }

        /// <summary>
        /// Writes all pending changes to the current file path — an explicit checkpoint.
        /// (Changes are also persisted automatically when the tool is disposed, so the file
        /// on disk always reflects the final session state.)
        /// Before saving, creates a numbered backup of the existing file (.001.bak, .002.bak, ...)
        /// so the original state can be restored later via <see cref="Restore"/>.
        /// </summary>
        /// <returns>A message describing the result: the backup file name, or "No changes to save" if nothing loaded.</returns>
        public string Save()
        {
            if (_workbook == null) return "No changes to save — no workbook is open.";

            var backupName = CreateBackup(_filePath);
            _workbook.Save(_filePath);
            PatchChartCaches(_filePath);
            Log.LogStep($"SpreadsheetTool.Save: saved to '{_filePath}', backup='{backupName}'");
            return string.IsNullOrEmpty(backupName)
                ? $"Workbook saved to '{Path.GetFileName(_filePath)}'. (New file, no backup needed.)"
                : $"Workbook saved to '{Path.GetFileName(_filePath)}'. The previous version was backed up as '{backupName}'.";
        }

        /// <summary>
        /// Writes all pending changes to a new file path.
        /// Subsequent Save() calls will use the new path.
        /// If the target file already exists, a numbered backup is created first.
        /// </summary>
        /// <param name="newFilePath">Path for the new .xlsx file, Unix style relative to the workspace root (e.g. "/folder/file.xlsx").</param>
        /// <returns>A message describing the result.</returns>
        public string SaveAs(string newFilePath)
        {
            if (_workbook == null) return "No changes to save — no workbook is open.";

            var resolved = SandboxPath.Resolve(newFilePath);
            var backupName = File.Exists(resolved) ? CreateBackup(resolved) : null;
            _workbook.Save(resolved);
            PatchChartCaches(resolved);
            var oldPath = Path.GetFileName(_filePath);
            _filePath = resolved;
            Log.LogStep($"SpreadsheetTool.SaveAs: saved to '{_filePath}'");
            return backupName != null
                ? $"Workbook saved as '{Path.GetFileName(resolved)}'. The existing file at that path was backed up as '{backupName}'."
                : $"Workbook saved as '{Path.GetFileName(resolved)}'.";
        }

        /// <summary>
        /// Creates a numbered backup of the specified file.
        /// Backup files follow the pattern: filename.NNN.bak (e.g. "data.001.bak", "data.002.bak").
        /// Never overwrites existing backups.
        /// </summary>
        /// <returns>The backup file name (without directory), or null if no file existed to back up.</returns>
        private static string? CreateBackup(string filePath)
        {
            if (!File.Exists(filePath)) return null;

            var dir = Path.GetDirectoryName(filePath) ?? ".";
            var nameWithoutExt = Path.GetFileNameWithoutExtension(filePath);

            // Find the next available backup number
            for (int i = 1; i <= 9999; i++)
            {
                var backupName = $"{nameWithoutExt}.{i:D3}.bak";
                var backupPath = Path.Combine(dir, backupName);
                if (!File.Exists(backupPath))
                {
                    File.Copy(filePath, backupPath);
                    return backupName;
                }
            }

            // Fallback: extremely unlikely, use timestamp
            var ts = DateTime.Now.ToString("yyyyMMddHHmmss");
            var fallbackName = $"{nameWithoutExt}.{ts}.bak";
            File.Copy(filePath, Path.Combine(dir, fallbackName));
            return fallbackName;
        }

        /// <summary>
        /// Restores the workbook to its state from the most recent backup (.bak file).
        /// The backup with the highest number is restored (most recent save point).
        /// The current (modified) workbook is replaced with the backup copy, and the
        /// backup file is preserved (not deleted) for future rollbacks.
        /// </summary>
        /// <returns>A message describing the restore result.</returns>
        public string Restore()
        {
            if (_workbook == null) return "No workbook is open. Nothing to restore.";

            var dir = Path.GetDirectoryName(_filePath) ?? ".";
            var nameWithoutExt = Path.GetFileNameWithoutExtension(_filePath);

            // Find the most recent backup (highest number)
            var backupFiles = Directory.GetFiles(dir, $"{nameWithoutExt}.*.bak")
                .OrderByDescending(f => f)
                .ToList();

            if (backupFiles.Count == 0)
                return "No backup file found. The workbook was never saved with backup enabled.";

            var latestBackup = backupFiles[0];
            var backupName = Path.GetFileName(latestBackup);

            try
            {
                _workbook.Dispose();
                File.Copy(latestBackup, _filePath, overwrite: true);
                _workbook = new Workbook(_filePath);
                Log.LogStep($"SpreadsheetTool.Restore: restored '{_filePath}' from '{backupName}'");
                return $"Workbook restored from backup '{backupName}'. The backup file has been preserved.";
            }
            catch (Exception ex)
            {
                Log.LogStep($"SpreadsheetTool.Restore: FAILED — {ex.Message}");
                return $"Restore failed: {ex.Message}";
            }
        }

        /// <summary>
        /// Explicit interface implementation — NOT an agent tool (the orchestrator disposes
        /// agents automatically when the loop ends). Persists any unsaved changes first so the
        /// file on disk always reflects the final session state, then releases the workbook.
        /// </summary>
        void IDisposable.Dispose()
        {
            try
            {
                if (_workbook != null && !string.IsNullOrEmpty(_filePath))
                {
                    var backupName = CreateBackup(_filePath);
                    _workbook.Save(_filePath);
                    PatchChartCaches(_filePath);
                    Log.LogStep(string.IsNullOrEmpty(backupName)
                        ? $"SpreadsheetTool.Dispose: auto-saved '{_filePath}' (new file, no backup)"
                        : $"SpreadsheetTool.Dispose: auto-saved '{_filePath}' (backup '{backupName}')");
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
        /// <returns>True if the sheet was renamed.</returns>
        public bool RenameWorksheet(string currentName, string newName)
        {
            var ws = FindSheet(currentName);
            if (ws == null) return false;
            ws.Name = newName;
            return true;
        }

        /// <summary>
        /// Shows or hides gridlines on a worksheet.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="show">True to show gridlines, false to hide them.</param>
        /// <returns>True if the setting was applied.</returns>
        public bool ShowGridlines(string sheetName, bool show)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.ShowGridlines = show;
            return true;
        }

        /// <summary>
        /// Shows or hides row and column headers (the gray 1,2,3... / A,B,C... area) on a worksheet.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="show">True to show headers, false to hide them.</param>
        /// <returns>True if the setting was applied.</returns>
        public bool ShowRowColumnHeaders(string sheetName, bool show)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.ShowRowColumnHeaders = show;
            return true;
        }

        /// <summary>
        /// Sets the zoom percentage for a worksheet (10-400).
        /// 100 = normal zoom.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="zoomPercentage">Zoom level between 10 and 400.</param>
        /// <returns>True if zoom was set.</returns>
        public bool SetZoom(string sheetName, int zoomPercentage)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.Zoom = zoomPercentage;
            return true;
        }

        /// <summary>
        /// Protects a worksheet so its structure cannot be modified.
        /// After protection, cells marked as locked (true by default) become read-only.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <returns>True if the sheet was protected.</returns>
        public bool ProtectSheet(string sheetName)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.Protect();
            return true;
        }

        /// <summary>
        /// Removes protection from a previously protected worksheet,
        /// allowing edits to locked cells again.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <returns>True if the sheet was unprotected.</returns>
        public bool UnprotectSheet(string sheetName)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.Unprotect();
            return true;
        }

        /// <summary>Adds a new, empty worksheet to the workbook.</summary>
        /// <param name="name">Name for the new worksheet (e.g. "Riepilogo"). Must be unique in the workbook.</param>
        /// <returns>True if the worksheet was added.</returns>
        public bool AddWorksheet(string name)
        {
            try
            {
                _workbook.Worksheets.Add(name);
                Log.LogStep($"SpreadsheetTool.AddWorksheet: added '{name}'");
                return true;
            }
            catch (Exception ex)
            {
                Log.LogStep($"SpreadsheetTool.AddWorksheet: FAILED — {ex.Message}");
                return false;
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
            return ws?.Cells[cellReference]?.DisplayStringValue;
        }

        /// <summary>
        /// Sets the value of a cell. Auto-detects numbers, booleans, dates, and strings.
        /// </summary>
        /// <param name="sheetName">Worksheet name (case-sensitive, from GetSheetNames()).</param>
        /// <param name="cellReference">Cell reference in A1 notation (e.g. "A1", "C5").</param>
        /// <param name="value">The value to write. Parsed automatically.</param>
        /// <returns>True if the cell was updated.</returns>
        public bool SetCellValue(string sheetName, string cellReference, string value)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            var (_, col) = ParseCellRef(cellReference);
            var fresh = IsFreshColumn(ws, col);
            SetCellValueAuto(ws.Cells[cellReference], value);
            if (fresh) ApplyAutoWidth(ws, col, new[] { value }); // first write to a fresh column → auto width
            Log.LogStep($"SpreadsheetTool.SetCellValue: '{sheetName}'!{cellReference} = '{value}'");
            return true;
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
            if (ws == null) return null;
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
        /// <returns>True if the formula was set.</returns>
        public bool SetCellFormula(string sheetName, string cellReference, string formula)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            var cell = ws.Cells[cellReference];
            if (cell == null) return false;
            var (_, col) = ParseCellRef(cellReference);
            var fresh = IsFreshColumn(ws, col);
            cell.Formula = formula;
            if (fresh) ApplyAutoWidth(ws, col, new[] { formula });
            return true;
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
            var cell = ws?.Cells[cellReference];
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
        /// <returns>2D string array [row][col], or null if the sheet is not found.</returns>
        public string[][]? GetRange(string sheetName, string startCell, string endCell)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return null;

            var (startRow, startCol) = ParseCellRef(startCell);
            var (endRow, endCol) = ParseCellRef(endCell);

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

        /// <summary>
        /// Writes a 2D string array starting at the specified cell.
        /// Auto-detects number, boolean, date, and string values.
        /// </summary>
        /// <param name="sheetName">Worksheet name (case-sensitive, from GetSheetNames()).</param>
        /// <param name="startCell">Top-left cell (e.g. "A1").</param>
        /// <param name="values">2D array of STRINGS [row][col]; pass numbers as strings (e.g. "14500") — the tool auto-detects and stores them as numbers. Rows may have different lengths.</param>
        /// <returns>True if the data was written.</returns>
        public bool SetRange(string sheetName, string startCell, string[][] values)
        {
            var ws = FindSheet(sheetName);
            if (ws == null || values.Length == 0) return false;

            var (startRow, startCol) = ParseCellRef(startCell);

            // Fresh columns (first use in this session, no prior content) get an auto width.
            var freshCols = new List<(int Col, List<string> Written)>();
            for (int c = 0; c < values[0].Length; c++)
            {
                if (!IsFreshColumn(ws, startCol + c)) continue;
                var written = new List<string>();
                for (int r = 0; r < values.Length && c < values[r].Length; r++)
                    written.Add(values[r][c]);
                freshCols.Add((startCol + c, written));
            }

            for (int r = 0; r < values.Length; r++)
            {
                for (int c = 0; c < values[r].Length; c++)
                {
                    var cell = ws.Cells[startRow + r, startCol + c];
                    if (cell != null)
                        SetCellValueAuto(cell, values[r][c]);
                }
            }

            foreach (var (col, written) in freshCols)
                ApplyAutoWidth(ws, col, written);

            Log.LogStep($"SpreadsheetTool.SetRange: '{sheetName}'!{startCell} ({values.Length} rows)");
            return true;
        }

        /// <summary>
        /// Appends rows after the last used row on the worksheet.
        /// Scans column A for the first empty cell, then writes from there.
        /// </summary>
        /// <param name="sheetName">Worksheet name (case-sensitive, from GetSheetNames()).</param>
        /// <param name="rows">Array of rows to append.</param>
        /// <returns>True if the rows were appended.</returns>
        public bool AppendRows(string sheetName, string[][] rows)
        {
            var ws = FindSheet(sheetName);
            if (ws == null || rows.Length == 0) return false;

            int startRow = 0;
            while (true)
            {
                var cell = ws.Cells[startRow, 0];
                if (cell == null || string.IsNullOrEmpty(cell.DisplayStringValue))
                    break;
                startRow++;
            }

            // Auto width for columns that had no content before this append.
            var freshCols = new List<(int Col, List<string> Written)>();
            for (int c = 0; c < rows[0].Length; c++)
            {
                if (!IsFreshColumn(ws, c)) continue;
                var written = new List<string>();
                for (int r = 0; r < rows.Length && c < rows[r].Length; r++)
                    written.Add(rows[r][c]);
                freshCols.Add((c, written));
            }

            for (int r = 0; r < rows.Length; r++)
            {
                for (int c = 0; c < rows[r].Length; c++)
                {
                    var cell = ws.Cells[startRow + r, c];
                    if (cell != null)
                        SetCellValueAuto(cell, rows[r][c]);
                }
            }

            foreach (var (col, written) in freshCols)
                ApplyAutoWidth(ws, col, written);

            Log.LogStep($"SpreadsheetTool.AppendRows: '{sheetName}' ({rows.Length} rows from row {startRow})");
            return true;
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
        /// <returns>True if the cells were merged.</returns>
        public bool MergeCells(string sheetName, string range)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;

            var (startRef, endRef) = ParseRange(range);
            var (startRow, startCol) = ParseCellRef(startRef);
            var (endRow, endCol) = ParseCellRef(endRef);

            int totalRows = endRow - startRow + 1;
            int totalCols = endCol - startCol + 1;
            ws.Cells.Merge(startRow, startCol, totalRows, totalCols);
            Log.LogStep($"SpreadsheetTool.MergeCells: '{sheetName}'!{range}");
            return true;
        }

        // ──────────────────────────────────────────────
        //  Style — font
        // ──────────────────────────────────────────────

        /// <summary>
        /// Sets font properties (name, size, bold, italic, color) on one or more cells.
        /// Specify a cell (e.g. "A1") or range (e.g. "A1:C10").
        /// Colors use "#RRGGBB" hex format (e.g. "#FF0000" for red).
        /// Pass null / 0 to leave a property unchanged.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="cellOrRange">Cell reference or range (e.g. "A1" or "A1:C10").</param>
        /// <param name="fontName">Font family (e.g. "Calibri", "Arial").</param>
        /// <param name="fontSize">Size in points (e.g. 11).</param>
        /// <param name="bold">True = bold.</param>
        /// <param name="italic">True = italic.</param>
        /// <param name="fontColorHex">Font color "#RRGGBB".</param>
        /// <returns>True if the font was applied.</returns>
        public bool SetCellFont(string sheetName, string cellOrRange,
            string? fontName = null, double fontSize = 0,
            bool? bold = null, bool? italic = null, string? fontColorHex = null)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            var cells = ResolveCells(ws, cellOrRange);
            if (cells == null) return false;

            foreach (var cell in cells)
            {
                if (cell == null) continue;
                var style = cell.GetStyle();
                if (!string.IsNullOrEmpty(fontName)) style.Font.Name = fontName;
                if (fontSize > 0) style.Font.Size = fontSize;
                if (bold.HasValue) style.Font.IsBold = bold.Value;
                if (italic.HasValue) style.Font.IsItalic = italic.Value;
                if (!string.IsNullOrEmpty(fontColorHex)) style.Font.Color = ParseColor(fontColorHex);
                cell.SetStyle(style);
            }
            return true;
        }

        // ──────────────────────────────────────────────
        //  Style — fill (background color)
        // ──────────────────────────────────────────────

        /// <summary>
        /// Sets the background (fill) color of one or more cells.
        /// Pass null or "none" to remove the fill.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="cellOrRange">Cell or range (e.g. "A1" or "A1:C10").</param>
        /// <param name="colorHex">Background color "#RRGGBB" (e.g. "#FFFF00" for yellow). Null or "none" clears fill.</param>
        /// <returns>True if the fill was applied.</returns>
        public bool SetCellFill(string sheetName, string cellOrRange, string? colorHex = null)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            var cells = ResolveCells(ws, cellOrRange);
            if (cells == null) return false;

            foreach (var cell in cells)
            {
                if (cell == null) continue;
                var style = cell.GetStyle();
                if (string.IsNullOrEmpty(colorHex) || colorHex.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    style.Pattern = FillPattern.None;
                }
                else
                {
                    style.Pattern = FillPattern.Solid;
                    style.ForegroundColor = ParseColor(colorHex);
                }
                cell.SetStyle(style);
            }
            return true;
        }

        // ──────────────────────────────────────────────
        //  Style — alignment
        // ──────────────────────────────────────────────

        /// <summary>
        /// Sets horizontal/vertical alignment and text wrapping on one or more cells.
        /// Horizontal: "General", "Left", "Center", "Right", "Fill", "Justify", "CenterAcrossSelection", "Distributed"
        /// Vertical: "Top", "Center", "Bottom", "Justify", "Distributed"
        /// Pass null to leave a property unchanged.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="cellOrRange">Cell or range (e.g. "A1" or "A1:C10").</param>
        /// <param name="horizontalAlignment">Horizontal alignment value.</param>
        /// <param name="verticalAlignment">Vertical alignment value.</param>
        /// <param name="wrapText">True = wrap, false = no wrap.</param>
        /// <returns>True if alignment was applied.</returns>
        public bool SetCellAlignment(string sheetName, string cellOrRange,
            string? horizontalAlignment = null, string? verticalAlignment = null, bool? wrapText = null)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            var cells = ResolveCells(ws, cellOrRange);
            if (cells == null) return false;

            foreach (var cell in cells)
            {
                if (cell == null) continue;
                var style = cell.GetStyle();
                if (!string.IsNullOrEmpty(horizontalAlignment))
                    style.HorizontalAlignment = ParseHorizontalAlignment(horizontalAlignment);
                if (!string.IsNullOrEmpty(verticalAlignment))
                    style.VerticalAlignment = ParseVerticalAlignment(verticalAlignment);
                if (wrapText.HasValue)
                    style.WrapText = wrapText.Value;
                cell.SetStyle(style);
            }
            return true;
        }

        // ──────────────────────────────────────────────
        //  Style — number format
        // ──────────────────────────────────────────────

        /// <summary>
        /// Sets a custom number format code on one or more cells.
        /// Common codes:
        ///   "#,##0.00" — thousands + 2 decimals
        ///   "€ #,##0.00" — euro currency
        ///   "0%" — percentage
        ///   "0.00%" — percentage with 2 decimals
        ///   "dd/mm/yyyy" — date
        ///   "dd/mm/yyyy hh:mm" — date + time
        ///   "@" — text
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="cellOrRange">Cell or range (e.g. "A1" or "A1:C10").</param>
        /// <param name="formatCode">Custom number format code (e.g. "#,##0.00").</param>
        /// <returns>True if the format was applied.</returns>
        public bool SetCellNumberFormat(string sheetName, string cellOrRange, string formatCode)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            var cells = ResolveCells(ws, cellOrRange);
            if (cells == null) return false;

            foreach (var cell in cells)
            {
                if (cell == null) continue;
                var style = cell.GetStyle();
                style.Custom = formatCode;
                cell.SetStyle(style);
            }
            return true;
        }

        // ──────────────────────────────────────────────
        //  Style — borders
        // ──────────────────────────────────────────────

        /// <summary>
        /// Applies borders to one or more cells.
        /// Styles: "None", "Thin", "Medium", "Thick", "Dotted", "Dashed", "Double", "Hair"
        /// Sides: "All", "Outline", "Inside", "Top", "Bottom", "Left", "Right"
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="cellOrRange">Cell or range (e.g. "A1" or "A1:C10").</param>
        /// <param name="borderStyle">Line style (default "Thin").</param>
        /// <param name="borderSide">Which sides (default "All").</param>
        /// <param name="borderColorHex">Color "#RRGGBB" (default black).</param>
        /// <returns>True if borders were applied.</returns>
        public bool SetCellBorders(string sheetName, string cellOrRange,
            string borderStyle = "Thin", string borderSide = "All", string? borderColorHex = null)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            var cells = ResolveCells(ws, cellOrRange);
            if (cells == null) return false;

            var styleType = ParseBorderStyle(borderStyle);
            var color = !string.IsNullOrEmpty(borderColorHex)
                ? ParseColor(borderColorHex)
                : Color.FromArgb(255, 0, 0, 0);
            var sides = ParseBorderSides(borderSide);

            foreach (var cell in cells)
            {
                if (cell == null) continue;
                var style = cell.GetStyle();
                var borders = style.Borders;

                foreach (var side in sides)
                {
                    switch (side)
                    {
                        case "Top":
                            borders.Top.LineStyle = styleType;
                            borders.Top.Color = color;
                            break;
                        case "Bottom":
                            borders.Bottom.LineStyle = styleType;
                            borders.Bottom.Color = color;
                            break;
                        case "Left":
                            borders.Left.LineStyle = styleType;
                            borders.Left.Color = color;
                            break;
                        case "Right":
                            borders.Right.LineStyle = styleType;
                            borders.Right.Color = color;
                            break;
                    }
                }
                cell.SetStyle(style);
            }
            return true;
        }

        // ──────────────────────────────────────────────
        //  Style — all-in-one
        // ──────────────────────────────────────────────

        /// <summary>
        /// Applies multiple style properties in a single call.
        /// Only non-null/non-default parameters are applied.
        /// Colors: "#RRGGBB". Horizontal: "Left","Center","Right". Vertical: "Top","Center","Bottom".
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="cellOrRange">Cell or range (e.g. "A1" or "A1:C10").</param>
        /// <param name="fontName">Font family.</param>
        /// <param name="fontSize">Font size in points.</param>
        /// <param name="bold">True = bold.</param>
        /// <param name="italic">True = italic.</param>
        /// <param name="fontColorHex">Font color "#RRGGBB".</param>
        /// <param name="fillColorHex">Background color "#RRGGBB".</param>
        /// <param name="horizontalAlignment">Horizontal: "Left","Center","Right".</param>
        /// <param name="verticalAlignment">Vertical: "Top","Center","Bottom".</param>
        /// <param name="wrapText">True = wrap text.</param>
        /// <param name="numberFormat">Custom number format (e.g. "#,##0.00").</param>
        /// <returns>True if the style was applied.</returns>
        public bool ApplyStyle(string sheetName, string cellOrRange,
            string? fontName = null, double fontSize = 0,
            bool? bold = null, bool? italic = null,
            string? fontColorHex = null, string? fillColorHex = null,
            string? horizontalAlignment = null, string? verticalAlignment = null,
            bool? wrapText = null, string? numberFormat = null)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            var cells = ResolveCells(ws, cellOrRange);
            if (cells == null) return false;

            foreach (var cell in cells)
            {
                if (cell == null) continue;
                var style = cell.GetStyle();

                if (!string.IsNullOrEmpty(fontName)) style.Font.Name = fontName;
                if (fontSize > 0) style.Font.Size = fontSize;
                if (bold.HasValue) style.Font.IsBold = bold.Value;
                if (italic.HasValue) style.Font.IsItalic = italic.Value;
                if (!string.IsNullOrEmpty(fontColorHex)) style.Font.Color = ParseColor(fontColorHex);

                if (!string.IsNullOrEmpty(fillColorHex))
                {
                    style.Pattern = FillPattern.Solid;
                    style.ForegroundColor = ParseColor(fillColorHex);
                }

                if (!string.IsNullOrEmpty(horizontalAlignment))
                    style.HorizontalAlignment = ParseHorizontalAlignment(horizontalAlignment);
                if (!string.IsNullOrEmpty(verticalAlignment))
                    style.VerticalAlignment = ParseVerticalAlignment(verticalAlignment);
                if (wrapText.HasValue)
                    style.WrapText = wrapText.Value;
                if (!string.IsNullOrEmpty(numberFormat))
                    style.Custom = numberFormat;

                cell.SetStyle(style);
            }
            return true;
        }

        // ──────────────────────────────────────────────
        //  Style — header row shortcut
        // ──────────────────────────────────────────────

        /// <summary>
        /// Applies a bold white-on-blue header style to the first row.
        /// Detects the used column count from row 0.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <returns>True if the header was formatted.</returns>
        public bool FormatHeaderRow(string sheetName)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;

            int maxCol = 0;
            while (true)
            {
                var cell = ws.Cells[0, maxCol];
                if (cell == null || string.IsNullOrEmpty(cell.DisplayStringValue))
                    break;
                maxCol++;
            }
            if (maxCol == 0) return false;

            for (int c = 0; c < maxCol; c++)
            {
                var cell = ws.Cells[0, c];
                if (cell == null) continue;
                var style = cell.GetStyle();
                style.Font.IsBold = true;
                style.Font.Color = Color.FromArgb(255, 255, 255, 255);
                style.Pattern = FillPattern.Solid;
                style.ForegroundColor = Color.FromArgb(255, 34, 120, 212);
                cell.SetStyle(style);
            }
            return true;
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
        /// <returns>True if the row height was set.</returns>
        public bool SetRowHeight(string sheetName, int rowIndex, double heightInPoints)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.Cells.Rows[rowIndex].Height = heightInPoints >= 0 ? heightInPoints : null;
            return true;
        }

        /// <summary>
        /// Hides a specific row (0-based index).
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="rowIndex">0-based row index.</param>
        /// <returns>True if the row was hidden.</returns>
        public bool HideRow(string sheetName, int rowIndex)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.Cells.Rows[rowIndex].IsHidden = true;
            return true;
        }

        /// <summary>
        /// Unhides a previously hidden row.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="rowIndex">0-based row index.</param>
        /// <returns>True if the row was unhidden.</returns>
        public bool UnhideRow(string sheetName, int rowIndex)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.Cells.Rows[rowIndex].IsHidden = false;
            return true;
        }

        /// <summary>
        /// Sets the width of a specific column (0-based index).
        /// Width is measured in characters of the default font.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="columnIndex">0-based column index.</param>
        /// <param name="widthInCharacters">Column width in character units.</param>
        /// <returns>True if the column width was set.</returns>
        public bool SetColumnWidth(string sheetName, int columnIndex, double widthInCharacters)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.Cells.Columns[columnIndex].Width = widthInCharacters;
            return true;
        }

        /// <summary>
        /// Hides a specific column (0-based index).
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="columnIndex">0-based column index.</param>
        /// <returns>True if the column was hidden.</returns>
        public bool HideColumn(string sheetName, int columnIndex)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.Cells.Columns[columnIndex].IsHidden = true;
            return true;
        }

        /// <summary>
        /// Unhides a previously hidden column.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="columnIndex">0-based column index.</param>
        /// <returns>True if the column was unhidden.</returns>
        public bool UnhideColumn(string sheetName, int columnIndex)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.Cells.Columns[columnIndex].IsHidden = false;
            return true;
        }

        // ──────────────────────────────────────────────
        //  Charts
        // ──────────────────────────────────────────────

        /// <summary>
        /// Adds a new chart to a worksheet and returns its index.
        /// The chart is placed in the specified cell-anchored rectangle.
        /// The data range is normalized (sheet name defaulted to the target sheet, "$" removed)
        /// and the chart data is populated deterministically (series + cached values), so the
        /// chart renders with data even in apps that do not refresh caches on open.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="chartType">Chart type: "Column", "Bar", "Line", "Pie", "Area", "Scatter", "Doughnut", "Radar", etc.</param>
        /// <param name="dataRange">Cell range for chart data (e.g. "Sheet1!$A$1:$B$5" or "A1:B5").</param>
        /// <param name="upperLeftRow">Zero-based row for upper-left anchor.</param>
        /// <param name="upperLeftColumn">Zero-based column for upper-left anchor.</param>
        /// <param name="lowerRightRow">Zero-based row for lower-right anchor.</param>
        /// <param name="lowerRightColumn">Zero-based column for lower-right anchor.</param>
        /// <returns>The zero-based chart index, or -1 on failure.</returns>
        public int AddChart(string sheetName, string chartType, string dataRange,
            int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return -1;
            var type = ParseChartType(chartType);
            var range = NormalizeChartRange(dataRange, ws.Name);
            var idx = ws.Charts.Add(type, range, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
            // The FOSS chart template saves an EMPTY numCache (ptCount=0), which renders as a
            // blank chart in LibreOffice/Excel. The cache is populated deterministically at save
            // time (PatchChartCaches) from the referenced cells, so the chart always shows data.
            Log.LogStep($"SpreadsheetTool.AddChart: '{sheetName}' {chartType} (index {idx}) range='{range}'");
            return idx;
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
            // Quote sheet names containing characters that break the range syntax.
            if (sheet.Any(ch => !char.IsLetterOrDigit(ch) && ch != '_'))
                sheet = "'" + sheet + "'";
            return $"{sheet}!{cells}";
        }

        /// <summary>
        /// Returns the number of charts on a worksheet.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <returns>Chart count, or -1 if sheet not found.</returns>
        public int GetChartsCount(string sheetName)
        {
            var ws = FindSheet(sheetName);
            return ws?.Charts.Count ?? -1;
        }

        /// <summary>
        /// Returns summary info about all charts on a worksheet as a 2D array.
        /// Columns: Index, Name, Type, Position.
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
        /// <returns>True if the hyperlink was added.</returns>
        public bool AddHyperlink(string sheetName, string cellName, int totalRows, int totalColumns, string address)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.Hyperlinks.Add(cellName, totalRows, totalColumns, address);
            return true;
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
        /// <returns>True if removed.</returns>
        public bool RemoveHyperlink(string sheetName, int index)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.Hyperlinks.RemoveAt(index);
            return true;
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
        /// <returns>True if the comment was added.</returns>
        public bool AddComment(string sheetName, string cellReference, string text, string? author = null)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            var comment = ws.Comments.Add(cellReference);
            comment.Note = text;
            if (!string.IsNullOrEmpty(author)) comment.Author = author;
            return true;
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
        /// <returns>True if removed.</returns>
        public bool RemoveComment(string sheetName, string cellReference)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.Comments.RemoveAt(cellReference);
            return true;
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
        /// <returns>True if the picture was added.</returns>
        public bool AddPicture(string sheetName, string imageFilePath,
            int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.Pictures.Add(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, imageFilePath);
            return true;
        }

        /// <summary>
        /// Removes a picture by its zero-based index.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="index">Zero-based picture index (0 to GetPicturesCount()-1).</param>
        /// <returns>True if removed.</returns>
        public bool RemovePicture(string sheetName, int index)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.Pictures.RemoveAt(index);
            return true;
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
        /// <returns>True if the table was added.</returns>
        public bool AddTable(string sheetName, string startCell, string endCell, bool hasHeaders = true)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.ListObjects.Add(startCell, endCell, hasHeaders);
            Log.LogStep($"SpreadsheetTool.AddTable: '{sheetName}'!{startCell}:{endCell}");
            return true;
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
        /// <returns>True if removed.</returns>
        public bool RemoveTable(string sheetName, int index)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.ListObjects.RemoveAt(index);
            return true;
        }

        // ──────────────────────────────────────────────
        //  AutoFilter
        // ──────────────────────────────────────────────

        /// <summary>
        /// Enables the AutoFilter (drop-down arrows) on a worksheet for the specified range.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <param name="range">Range in A1 notation (e.g. "A1:C10").</param>
        /// <returns>True if the AutoFilter was set.</returns>
        public bool SetAutoFilter(string sheetName, string range)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            var (startRef, endRef) = ParseRange(range);
            var (startRow, startCol) = ParseCellRef(startRef);
            var (endRow, endCol) = ParseCellRef(endRef);
            ws.AutoFilter.Range = range;
            return true;
        }

        /// <summary>
        /// Removes the AutoFilter from a worksheet.
        /// </summary>
        /// <param name="sheetName">Worksheet name (from GetSheetNames()).</param>
        /// <returns>True if the AutoFilter was removed.</returns>
        public bool RemoveAutoFilter(string sheetName)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
            ws.AutoFilter.Clear();
            return true;
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
        /// <returns>True if the defined name was added.</returns>
        public bool AddDefinedName(string name, string formula, string? localSheetName = null)
        {
            try
            {
                int? sheetIndex = null;
                if (!string.IsNullOrEmpty(localSheetName))
                {
                    var ws = FindSheet(localSheetName);
                    if (ws == null) return false;
                    // Find the sheet index
                    for (int i = 0; i < _workbook.Worksheets.Count; i++)
                    {
                        if (_workbook.Worksheets[i].Name == localSheetName)
                        { sheetIndex = i; break; }
                    }
                }
                _workbook.DefinedNames.Add(name, formula, sheetIndex);
                Log.LogStep($"SpreadsheetTool.AddDefinedName: '{name}' = {formula}");
                return true;
            }
            catch { return false; }
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
        /// <returns>True if removed.</returns>
        public bool RemoveDefinedName(string name)
        {
            try
            {
                for (int i = 0; i < _workbook.DefinedNames.Count; i++)
                {
                    if (_workbook.DefinedNames[i].Name == name)
                    {
                        _workbook.DefinedNames.RemoveAt(i);
                        return true;
                    }
                }
                return false;
            }
            catch { return false; }
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
        /// <returns>True if settings were applied.</returns>
        public bool SetPageSetup(string sheetName,
            string? orientation = null, string? paperSize = null,
            int? scale = null, int? fitToPagesWide = null, int? fitToPagesTall = null,
            string? printArea = null, bool? centerHorizontally = null, bool? centerVertically = null)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;
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

            return true;
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
        /// <returns>True if the conditional format was added.</returns>
        public bool AddConditionalFormat(string sheetName, string range,
            string conditionType, string? operatorType = null,
            string? formula1 = null, string? formula2 = null,
            string? fontColorHex = null, string? fillColorHex = null, bool? bold = null)
        {
            var ws = FindSheet(sheetName);
            if (ws == null) return false;

            var (startRef, endRef) = ParseRange(range);
            var (startRow, startCol) = ParseCellRef(startRef);
            var (endRow, endCol) = ParseCellRef(endRef);

            var area = new CellArea
            {
                StartRow = startRow,
                StartColumn = startCol,
                EndRow = endRow,
                EndColumn = endCol
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
            return true;
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
                return "{\"error\": \"No workbook loaded. Call Open(filePath) first.\"}";

            sampleRowCount = Math.Clamp(sampleRowCount, 1, 20);

            var result = new Dictionary<string, object?>
            {
                ["filePath"] = _filePath,
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

        /// <summary>True the first time this session writes to a column (a "fresh" column that
        /// had no prior content). Later writes keep the width that was already set.</summary>
        private bool IsFreshColumn(Worksheet ws, int col)
        {
            if (!_autoWidthCols.TryGetValue(ws.Name, out var set))
            {
                set = new HashSet<int>();
                _autoWidthCols[ws.Name] = set;
            }
            return set.Add(col);
        }

        /// <summary>Finds every column that already carries content (used when opening a file,
        /// so auto width never overrides an existing column layout).</summary>
        private static HashSet<int> ScanNonEmptyColumns(Worksheet ws)
        {
            var cols = new HashSet<int>();
            int lastUsedRow = -1;
            for (int r = 0; r < 5000; r++)
            {
                bool any = false;
                for (int c = 0; c < 500; c++)
                {
                    var cell = ws.Cells[r, c];
                    if (cell != null && (!string.IsNullOrEmpty(cell.DisplayStringValue) || !string.IsNullOrEmpty(cell.Formula)))
                    { cols.Add(c); any = true; }
                }
                if (any) lastUsedRow = r;
                else if (r > lastUsedRow + 5) break; // 5 empty rows → stop
            }
            return cols;
        }

        /// <summary>Content-based column width (the FOSS fork has no AutoFit API): the longest
        /// written value plus padding, bounded to a sane range.</summary>
        private static void ApplyAutoWidth(Worksheet ws, int col, IEnumerable<string> writtenValues)
        {
            var maxLen = writtenValues.Where(v => !string.IsNullOrEmpty(v)).Select(v => v.Length).DefaultIfEmpty(0).Max();
            ws.Cells.Columns[col].Width = Math.Clamp(maxLen + 2.0, 8.0, 60.0);
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

        private static readonly System.Text.RegularExpressions.Regex ChartCachePattern = new(
            @"<c:f>(?<range>[^<]+)</c:f>\s*<c:numCache><c:formatCode>[^<]*</c:formatCode><c:ptCount val=""0""/>",
            System.Text.RegularExpressions.RegexOptions.Compiled);

        /// <summary>Replaces every empty chart numCache with the actual values of the range
        /// referenced by the preceding &lt;c:f&gt;. Unchanged (empty) when no numeric value is found.</summary>
        private string PopulateChartCacheXml(string xml)
        {
            return ChartCachePattern.Replace(xml, m =>
            {
                var pts = BuildChartPointsXml(m.Groups["range"].Value);
                return pts != null
                    ? $"<c:f>{m.Groups["range"].Value}</c:f><c:numCache><c:formatCode>General</c:formatCode>{pts}</c:numCache>"
                    : m.Value;
            });
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
                var (startRef, endRef) = ParseRange(cells);
                var (r1, c1) = ParseCellRef(startRef);
                var (r2, c2) = ParseCellRef(endRef);
                r2 = Math.Min(r2, r1 + 1000);
                c2 = Math.Min(c2, c1 + 100);

                var vals = new List<string>();
                for (int r = r1; r <= r2; r++)
                    for (int c = c1; c <= c2; c++)
                    {
                        var cell = ws.Cells[r, c];
                        var s = cell?.StringValue;
                        if (string.IsNullOrEmpty(s)) continue;
                        if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
                            vals.Add(d.ToString(CultureInfo.InvariantCulture));
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

        private static List<Cell?> ResolveCells(Worksheet ws, string cellOrRange)
        {
            var cells = new List<Cell?>();

            if (cellOrRange.Contains(':'))
            {
                var parts = cellOrRange.Split(':');
                var (startRow, startCol) = ParseCellRef(parts[0]);
                var (endRow, endCol) = ParseCellRef(parts[1]);

                for (int r = startRow; r <= endRow; r++)
                    for (int c = startCol; c <= endCol; c++)
                        cells.Add(ws.Cells[r, c]);
            }
            else
            {
                cells.Add(ws.Cells[cellOrRange]);
            }

            return cells;
        }

        private static (int Row, int Col) ParseCellRef(string refStr)
        {
            int col = 0, i = 0;
            while (i < refStr.Length && char.IsLetter(refStr[i]))
            {
                col = col * 26 + (char.ToUpper(refStr[i]) - 'A' + 1);
                i++;
            }
            int row = int.Parse(refStr.AsSpan(i)) - 1;
            return (row, col - 1);
        }

        private static (string Start, string End) ParseRange(string range)
        {
            var parts = range.Split(':');
            return (parts[0], parts.Length > 1 ? parts[1] : parts[0]);
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
            if (int.TryParse(value, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var iVal))
            { cell.PutValue(iVal); return; }
            if (decimal.TryParse(value, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var dVal))
            { cell.PutValue(dVal); return; }
            if (DateTime.TryParse(value, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out var dtVal))
            { cell.PutValue(dtVal); return; }

            cell.PutValue(value);
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

        private static ChartType ParseChartType(string value)
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
                _ => ChartType.Column,
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
