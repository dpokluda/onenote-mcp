using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using OneNoteMcp.ComClient;
using OneNoteMcp.Configuration;

namespace OneNoteMcp.McpServer;

/// <summary>
/// The MCP server surface: the set of tools exposed to MCP hosts (Copilot CLI, Claude Desktop, etc.).
/// Each public method decorated with <see cref="McpServerToolAttribute"/> becomes a callable tool.
/// All OneNote access is delegated to <see cref="OneNoteComClient"/>; this type only handles
/// argument validation, category resolution, and JSON shaping of results.
/// </summary>
[McpServerToolType]
public sealed class OneNoteMcpServer
{
    private static readonly JsonSerializerOptions JsonOpts = new() { WriteIndented = true };

    private readonly OneNoteComClient _client;

    public OneNoteMcpServer(OneNoteComClient client)
    {
        _client = client;
    }

    // Maps an optional category argument to the categories to read: a single resolved category,
    // or all configured categories when the argument is omitted. Throws on an unknown name.
    private IReadOnlyList<OneNoteCategory> ResolveCategories(string? category)
    {
        if (string.IsNullOrWhiteSpace(category)) return _client.Categories;
        var resolved = _client.ResolveCategory(category);
        if (resolved is null)
        {
            var available = string.Join(", ", _client.Categories.Select(c => $"\"{c.Name}\""));
            throw new ArgumentException($"Unknown category '{category}'. Available categories: {available}.");
        }

        return new[] { resolved };
    }

    [McpServerTool(Name = "list_pages")]
    [Description("Lists pages for a category (e.g. \"TSG\" or \"On-Call Notes\"), reading across every section in that category, most recently modified first. Omit category to list across all categories. Returns page id, title, category, section, and modified time.")]
    public string ListPages(
        [Description("Category name to read, e.g. \"TSG\" or \"On-Call Notes\". Omit or leave empty to include all categories.")] string? category = null,
        [Description("Maximum number of pages to return (1-500). Default 50.")] int top = 50)
    {
        var pages = _client.ListPages(ResolveCategories(category), top);
        var result = pages.Select(p => new
        {
            id = p.Id,
            title = p.Title,
            category = p.Category,
            section = p.SectionDisplayName,
            lastModified = p.LastModifiedUtc,
            created = p.CreatedUtc,
        });
        return JsonSerializer.Serialize(result, JsonOpts);
    }

    [McpServerTool(Name = "search_pages")]
    [Description("Searches OneNote pages by title (case-insensitive substring match) within a category (e.g. \"TSG\" or \"On-Call Notes\"). Omit category to search all categories. For full-text content search, call list_pages then get_page on candidates.")]
    public string SearchPages(
        [Description("Search term to match against page titles.")] string query,
        [Description("Category name to search, e.g. \"TSG\" or \"On-Call Notes\". Omit or leave empty to search all categories.")] string? category = null,
        [Description("Maximum number of matches to return (1-500). Default 20.")] int top = 20)
    {
        if (string.IsNullOrWhiteSpace(query)) return "[]";
        var pages = _client.SearchPagesByTitle(ResolveCategories(category), query, top);
        var result = pages.Select(p => new
        {
            id = p.Id,
            title = p.Title,
            category = p.Category,
            section = p.SectionDisplayName,
            lastModified = p.LastModifiedUtc,
        });
        return JsonSerializer.Serialize(result, JsonOpts);
    }

    [McpServerTool(Name = "get_page")]
    [Description("Returns the content of a single OneNote page by id. Default format is plain text; pass format='markdown' for structured Markdown or format='xml' for raw OneNote XML.")]
    public string GetPage(
        [Description("OneNote page id as returned by list_pages or search_pages.")] string pageId,
        [Description("Output format: 'text' (default, recommended), 'markdown' (structured Markdown), or 'xml' (raw OneNote XML).")] string format = "text")
    {
        if (string.IsNullOrWhiteSpace(pageId))
        {
            throw new ArgumentException("pageId is required.", nameof(pageId));
        }

        return format?.Trim().ToLowerInvariant() switch
        {
            "xml" => _client.GetPageXml(pageId),
            "markdown" or "md" => _client.GetPageMarkdown(pageId),
            _ => _client.GetPageText(pageId),
        };
    }

    [McpServerTool(Name = "list_categories")]
    [Description("Lists the configured OneNote categories (e.g. \"TSG\", \"On-Call Notes\") and the sections each one reads from.")]
    public string ListCategories()
    {
        var categories = _client.Categories.Select(c => new
        {
            category = c.Name,
            sections = c.Sections.Select(s => new { section = s.EffectiveDisplayName, path = s.Path }),
        });
        return JsonSerializer.Serialize(new { backend = "OneNote desktop COM (local)", categories }, JsonOpts);
    }
}
