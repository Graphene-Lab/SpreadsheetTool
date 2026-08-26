# SpreadsheetTool — Spreadsheet (XLSX) agent tool

Spreadsheet operations for LLM agents: open/create, cells, ranges, styles, charts,
tables. Implements `AIOrchestrator.API.IAgentTool`.

## What it is

A **plugin** for hosts built on [Graphene.AIOrchestrator](https://github.com/Graphene-Lab/AgentHarness)
(e.g. AIOffice): the assembly is loaded dynamically from the host's `Tools/` folder
at startup (or hot-added via a filesystem watcher) and its public methods become
LLM-callable tools through the `UISupportGeneric.Analyzer.GeToolDefinitions`
reflection pipeline.

## Usage

- **Host with dynamic loading** (AIOffice): drop `SpreadsheetTool.dll` +
  `SpreadsheetTool.xml` into the host's `Tools/` directory. The host scans the
  folder at startup and on new files (30 s debounce).
- **Host with static reference**: `ProjectReference` the project (or
  `PackageReference Graphene.SpreadsheetTool`) and register
  `typeof(AIOrchestrator.API.SpreadsheetTool)` in the agent types array passed to
  `AgentHarness.ExecuteAction`.

## Build

```
dotnet build -c Release
```

Compiled as **AnyCPU, RID-neutral**: the same `.dll` runs on Linux, Windows, macOS
and iOS. The XML documentation file (`SpreadsheetTool.xml`) is generated and must
ship next to the assembly — `UISupportGeneric` reads method/parameter descriptions
from it.

## NuGet

`Graphene.SpreadsheetTool`, date-versioned `1.yy.MM.dd`, published automatically on
every `master` push (`.github/workflows/publish.yml`, `NUGET_API_KEY` repo secret).
Local pack pushes with the `NuGetApiKey` env var; skip with `-p:SkipNuGetPush=true`.

## Notes

- **Charts** are written as standard OOXML with the data embedded (series, categories
  and cached values), so they render in Excel, Office Online and other spreadsheet
  viewers without any recalculation. Known limitation: **LibreOffice** may render the
  chart frame and axes but not the series for these charts (an import/rendering quirk
  of LibreOffice itself — the same files render correctly elsewhere).

## License

See [LICENSE.md](LICENSE.md).
