# OneNote MCP Server

An MCP server that exposes one or more OneNote sections, grouped into named **categories** (e.g. `TSG`, `On-Call Notes`), over the Model Context Protocol via `stdio`. Built on the official C# MCP SDK and the **OneNote desktop COM Interop API**, no Microsoft Graph, no AAD, no tenant consent dance.

## What it does

Reads pages from any number of OneNote sections, organized into named **categories**. A prompt references a category (e.g. *"check the TSGs"* or *"check on-call notes"*) and the server reads across **every section in that category**.

Out of the box it ships with two categories:

| Category | Sections |
|----------|----------|
| `TSG` | Governance Vteam Notebook > Policy > On-Call |
| `On-Call Notes` | Azure Policy Livesite Handoff > 2022–2026 Weekly Summaries |

Each section is resolved by hierarchical **path** (notebook / section group / section name), not by GUID, OneNote desktop assigns its own internal IDs that have nothing to do with the GUIDs in SharePoint URLs.

## APIs

| Tool | Purpose |
|------|---------|
| `list_categories` | Returns the configured categories and the sections each one reads from, plus backend info. |
| `list_pages` | Lists pages (most recently modified first), each tagged with its category and section. Params: `category` (optional, omit to include all categories), `top` (default 50). |
| `search_pages` | Substring match against page titles within a category. Params: `query`, `category` (optional, omit to search all categories), `top` (default 20). |
| `get_page` | Returns a single page as plain text (default), structured Markdown, or OneNote XML. Params: `pageId`, `format` (`text` \| `markdown` \| `xml`). |

Category names are matched case-insensitively and ignore spaces/punctuation, so `TSG`, `tsgs`, `on-call notes`, and `oncallnotes` all resolve.

## Prerequisites

1. **Windows** (COM Interop is Windows-only, the project targets `net10.0-windows`).
2. **.NET 10 SDK**.
3. **OneNote desktop (Office 2013 or later)** installed. The PIA is referenced directly from the GAC at:
   
   ```
   C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.OneNote.dll
   ```
4. **OneNote desktop must be running** with the target notebooks opened and visible in the notebooks pane. Every notebook referenced by a configured section path must be open. 
   
   First-time setup for each: open the notebook in OneNote on the web > click **Open in desktop app** > let it sync.

That's it. No `az login`, no Graph scopes, no app registration.

## Build

```powershell
dotnet build -c Release
```

Output DLL: `onenote-mcp\bin\Release\net10.0-windows\onenote-mcp.dll`

## Configuration

`onenote-mcp\appsettings.json` defines categories, each with one or more sections:

```json
{
  "OneNote": {
    "Categories": [
      {
        "Name": "TSG",
        "Sections": [
          { "Path": "Governance Vteam Notebook/Policy/On-Call", "DisplayName": "Policy / On-Call" }
        ]
      },
      {
        "Name": "On-Call Notes",
        "Sections": [
          { "Path": "Azure Policy Livesite Handoff/2026 Weekly Summaries", "DisplayName": "2026 Weekly Summaries" },
          { "Path": "Azure Policy Livesite Handoff/2025 Weekly Summaries", "DisplayName": "2025 Weekly Summaries" },
          { "Path": "Azure Policy Livesite Handoff/2024 Weekly Summaries", "DisplayName": "2024 Weekly Summaries" },
          { "Path": "Azure Policy Livesite Handoff/2023 Weekly Summaries", "DisplayName": "2023 Weekly Summaries" },
          { "Path": "Azure Policy Livesite Handoff/2022 Weekly Summaries", "DisplayName": "2022 Weekly Summaries" }
        ]
      }
    ]
  }
}
```

This mirrors the shipped `appsettings.json`. The two categories are:

| Category | Reads from |
|----------|------------|
| `TSG` | Troubleshooting guides: *Governance Vteam Notebook > Policy > On-Call* |
| `On-Call Notes` | Weekly on-call handoff summaries: *Azure Policy Livesite Handoff > 2022–2026 Weekly Summaries* |


- `Name`: how the category is selected from a prompt (matched case-insensitively, ignoring spaces/punctuation).
- `Sections[].Path`: forward-slash-separated names walking from the notebook down through any section groups to the section itself. Names must match exactly (case-insensitive).
- `Sections[].DisplayName`: optional label shown in tool output; defaults to the last path segment.

Add sections to a category, or add whole new categories, by editing this list. A section whose path can't be resolved at runtime is skipped with a warning rather than failing the whole call.

> **Note:** `appsettings.json` is bound as the `OneNote` configuration section, so deeply nested list overrides via environment variables are impractical — edit the JSON file to change categories.

## Wiring into Copilot CLI

Add an entry to `~/.copilot/mcp-config.json`:

```json
{
  "mcpServers": {
    "onenote-mcp": {
      "type": "local",
      "command": "dotnet",
      "args": [
        "E:\\projects\\onenote-mcp\\bin\\Release\\net10.0-windows\\onenote-mcp.dll"
      ],
      "tools": ["*"]
    }
  }
}
```

Then in Copilot CLI:

1. Run `/mcp` and confirm `onenote-mcp` shows as connected.
2. Call `list_categories` to verify the categories and sections resolve.
3. Call `list_pages` with `category="TSG"` and `top=10`.

The same config works in Claude Desktop or any other stdio MCP host.

## Operational notes

- **Restart after rebuilds.** Copilot CLI caches the MCP server process across tool calls. After rebuilding, run `/restart` (or close and reopen the host) — otherwise you'll keep hitting the old in-memory copy. Builds will also fail with file-lock errors while a host is alive; kill the locking `dotnet` process first.
- **Logs go to stderr.** stdout is reserved for JSON-RPC framing. Run the DLL manually with `dotnet path\to\OneNoteMcp.dll` to watch startup logs.
- **`appsettings.json` is loaded from the DLL folder**, not the caller's CWD. If you edit it, edit the copy in `bin\Release\net10.0-windows\`.

## Limitations

- **Title-only search.** `search_pages` does substring matching on titles. Body-text search would require fetching page content for each candidate — not implemented.
- **Read-only.** No tools to create/update/delete pages or sections.
- **Requires OneNote desktop to be running.** If OneNote isn't running or a referenced notebook isn't loaded, that section is skipped (or COM calls throw if no hierarchy is available at all).

## Troubleshooting

| Symptom | Likely cause |
|---------|--------------|
| `list_categories` shows old output | Copilot CLI is using the cached server process. Run `/restart`. |
| A category returns fewer sections than expected | A section path didn't resolve (notebook not open, or name mismatch). Check stderr logs for skip warnings and verify the section in OneNote's notebook pane. |
| `COMException 0x800401F0` (CoInitialize) | OneNote desktop isn't installed or running. Start OneNote. |
| Build fails with file lock on `OneNoteMcp.dll` | A live Copilot/Claude session has the DLL loaded. Find the PID and stop it: `Get-Process dotnet \| Where-Object { $_.Modules.FileName -contains "<dll path>" }`. |
