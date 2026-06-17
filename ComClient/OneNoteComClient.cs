using System.Globalization;
using System.Net;
using System.Text;
using System.Xml.Linq;
using HtmlAgilityPack;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Office.Interop.OneNote;
using OneNoteMcp.Configuration;

namespace OneNoteMcp.ComClient;

/// <summary>
/// Reads pages from the local OneNote desktop client via COM automation (PIA-bound).
/// Requires OneNote 2016+ installed and the target notebook already opened on this machine.
/// Identifies the target section by its display path (notebook/group/.../section) because
/// OneNote desktop assigns its own internal IDs unrelated to SharePoint URL GUIDs.
/// </summary>
public sealed class OneNoteComClient
{
    private static readonly XNamespace One = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    private readonly OneNoteOptions _options;
    private readonly ILogger<OneNoteComClient> _logger;

    // The OneNote COM Application is single-threaded; serialize every call through this lock.
    private readonly object _comLock = new();
    private Application? _app;

    public OneNoteComClient(IOptions<OneNoteOptions> options, ILogger<OneNoteComClient> logger)
    {
        _options = options.Value;
        _logger = logger;
    }

    /// <summary>All configured categories, in declaration order.</summary>
    public IReadOnlyList<OneNoteCategory> Categories => _options.Categories;

    /// <summary>Resolves a category by name: exact (normalized) match, else prefix match either direction.</summary>
    public OneNoteCategory? ResolveCategory(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return null;
        var target = Normalize(name);
        return _options.Categories.FirstOrDefault(c => Normalize(c.Name) == target)
            ?? _options.Categories.FirstOrDefault(c =>
            {
                var n = Normalize(c.Name);
                return n.StartsWith(target, StringComparison.Ordinal) || target.StartsWith(n, StringComparison.Ordinal);
            });
    }

    // Strips spaces/punctuation and lowercases so "On-Call Notes" matches "oncallnotes".
    private static string Normalize(string value) =>
        new string(value.Where(char.IsLetterOrDigit).ToArray()).ToLowerInvariant();

    // Lazily connects to the running OneNote desktop instance, reusing the COM object thereafter.
    private Application App
    {
        get
        {
            if (_app is not null) return _app;
            lock (_comLock)
            {
                if (_app is not null) return _app;
                _app = new Application();
                _logger.LogInformation("Connected to OneNote desktop via COM.");
                return _app;
            }
        }
    }

    /// <summary>
    /// Returns the most-recently-modified pages across every section of the given categories,
    /// each page tagged with its originating category and section. Unresolvable sections are
    /// skipped with a warning rather than failing the whole call.
    /// </summary>
    public IReadOnlyList<OneNotePageInfo> ListPages(IReadOnlyList<OneNoteCategory> categories, int top)
    {
        var doc = LoadHierarchy();
        var pages = new List<OneNotePageInfo>();
        foreach (var category in categories)
        {
            foreach (var sectionRef in category.Sections)
            {
                var section = ResolveSectionByPath(doc, sectionRef.Path);
                if (section is null)
                {
                    _logger.LogWarning(
                        "Section '{Path}' (category '{Category}') not found in OneNote hierarchy; skipping.",
                        sectionRef.Path, category.Name);
                    continue;
                }

                pages.AddRange(section.Elements(One + "Page").Select(p => new OneNotePageInfo(
                    Id: (string)p.Attribute("ID")!,
                    Title: (string?)p.Attribute("name") ?? "(untitled)",
                    LastModifiedUtc: ParseTime(p.Attribute("lastModifiedTime")?.Value),
                    CreatedUtc: ParseTime(p.Attribute("dateTime")?.Value),
                    Category: category.Name,
                    SectionDisplayName: sectionRef.EffectiveDisplayName,
                    SectionPath: sectionRef.Path)));
            }
        }

        return pages
            .OrderByDescending(p => p.LastModifiedUtc)
            .Take(Math.Clamp(top, 1, 500))
            .ToList();
    }

    /// <summary>Case-insensitive substring match against page titles within the given categories.</summary>
    public IReadOnlyList<OneNotePageInfo> SearchPagesByTitle(IReadOnlyList<OneNoteCategory> categories, string query, int top)
    {
        return ListPages(categories, 500)
            .Where(p => p.Title.Contains(query, StringComparison.OrdinalIgnoreCase))
            .Take(Math.Clamp(top, 1, 500))
            .ToList();
    }

    /// <summary>Returns the raw OneNote XML for a single page.</summary>
    public string GetPageXml(string pageId)
    {
        lock (_comLock)
        {
            App.GetPageContent(pageId, out var xml, PageInfo.piBasic);
            return xml;
        }
    }

    /// <summary>Returns a single page flattened to plain text (title first, then body lines).</summary>
    public string GetPageText(string pageId)
    {
        var xml = GetPageXml(pageId);
        var doc = XDocument.Parse(xml);
        var title = doc.Descendants(One + "Title").Descendants(One + "OE").Descendants(One + "T").FirstOrDefault()?.Value;
        var body = doc.Descendants(One + "T")
            .Where(t => t.Ancestors(One + "Title").FirstOrDefault() is null)
            .Select(t => HtmlToPlainText(t.Value))
            .Where(s => !string.IsNullOrWhiteSpace(s));

        var lines = new List<string>();
        if (!string.IsNullOrWhiteSpace(title)) lines.Add(HtmlToPlainText(title));
        lines.AddRange(body);
        return string.Join("\n", lines);
    }

    /// <summary>Returns a single page rendered as Markdown (headings, lists, and tables preserved).</summary>
    public string GetPageMarkdown(string pageId)
    {
        var xml = GetPageXml(pageId);
        var doc = XDocument.Parse(xml);

        // OneNote references paragraph styles (h1..h6, normal, etc.) by index; build index -> name lookup.
        var styleMap = doc.Descendants(One + "QuickStyleDef")
            .GroupBy(d => (string?)d.Attribute("index") ?? string.Empty)
            .ToDictionary(g => g.Key, g => (string?)g.First().Attribute("name") ?? string.Empty);

        var sb = new StringBuilder();

        var title = doc.Descendants(One + "Title")
            .Descendants(One + "OE")
            .Descendants(One + "T")
            .FirstOrDefault()?.Value;
        if (!string.IsNullOrWhiteSpace(title))
        {
            sb.Append("# ").Append(ToInlineBreaks(InlineHtmlToMarkdown(title).Trim())).Append("\n\n");
        }

        foreach (var outline in doc.Descendants(One + "Outline"))
        {
            foreach (var children in outline.Elements(One + "OEChildren"))
            {
                RenderOeChildren(children, styleMap, sb, 0);
            }
        }

        return sb.ToString().TrimEnd() + "\n";
    }

    private void RenderOeChildren(XElement oeChildren, IReadOnlyDictionary<string, string> styleMap, StringBuilder sb, int depth)
    {
        foreach (var oe in oeChildren.Elements(One + "OE"))
        {
            RenderOe(oe, styleMap, sb, depth);
        }
    }

    private void RenderOe(XElement oe, IReadOnlyDictionary<string, string> styleMap, StringBuilder sb, int depth)
    {
        var table = oe.Element(One + "Table");
        if (table is not null)
        {
            RenderTable(table, sb);
        }
        else
        {
            var text = string.Concat(oe.Elements(One + "T").Select(t => InlineHtmlToMarkdown(t.Value))).Trim();
            var list = oe.Element(One + "List");
            var styleIndex = (string?)oe.Attribute("quickStyleIndex");
            var styleName = styleIndex is not null && styleMap.TryGetValue(styleIndex, out var n) ? n : string.Empty;

            if (list is not null)
            {
                var indent = new string(' ', depth * 2);
                var marker = list.Element(One + "Number") is not null ? "1." : "-";
                sb.Append(indent).Append(marker).Append(' ').Append(ToInlineBreaks(text)).Append('\n');
            }
            else if (IsHeading(styleName, out var level))
            {
                sb.Append(new string('#', level)).Append(' ').Append(ToInlineBreaks(text)).Append("\n\n");
            }
            else if (!string.IsNullOrEmpty(text))
            {
                sb.Append(ToParagraphBreaks(text)).Append("\n\n");
            }
        }

        var nextDepth = oe.Element(One + "List") is not null ? depth + 1 : depth;
        foreach (var children in oe.Elements(One + "OEChildren"))
        {
            RenderOeChildren(children, styleMap, sb, nextDepth);
        }
    }

    private void RenderTable(XElement table, StringBuilder sb)
    {
        var rows = table.Elements(One + "Row").ToList();
        if (rows.Count == 0) return;

        for (var i = 0; i < rows.Count; i++)
        {
            var cells = rows[i].Elements(One + "Cell").Select(GetCellText).ToList();
            sb.Append("| ").Append(string.Join(" | ", cells)).Append(" |\n");
            if (i == 0)
            {
                sb.Append("| ").Append(string.Join(" | ", cells.Select(_ => "---"))).Append(" |\n");
            }
        }
        sb.Append('\n');
    }

    private string GetCellText(XElement cell)
    {
        // Each OE inside a cell is its own line; preserve those breaks plus any inline <br>.
        var lines = cell.Descendants(One + "OE")
            .Select(oe => string.Concat(oe.Elements(One + "T").Select(t => InlineHtmlToMarkdown(t.Value))).Trim())
            .Where(s => s.Length > 0);
        return ToInlineBreaks(string.Join("\n", lines));
    }

    // Markdown paragraph hard break: two trailing spaces before the newline force a real line break.
    private static string ToParagraphBreaks(string text) => NormalizeNewlines(text).Replace("\n", "  \n");

    // Line breaks that must stay on a single logical line (table cells, headings, list items) become <br>.
    private static string ToInlineBreaks(string text) => NormalizeNewlines(text).Replace("\n", "<br>");

    private static string NormalizeNewlines(string text) => text.Replace("\r\n", "\n").Replace('\r', '\n');

    // OneNote heading styles are named "h1".."h6"; shift by one so the page title keeps the single '#'.
    private static bool IsHeading(string styleName, out int level)
    {
        level = 0;
        if (styleName.Length == 2 &&
            (styleName[0] == 'h' || styleName[0] == 'H') &&
            char.IsDigit(styleName[1]))
        {
            var parsed = styleName[1] - '0';
            if (parsed is >= 1 and <= 6)
            {
                level = Math.Min(parsed + 1, 6);
                return true;
            }
        }
        return false;
    }

    /// <summary>Converts a OneNote run's inline HTML (bold/italic spans, links, breaks) into Markdown.</summary>
    public static string InlineHtmlToMarkdown(string html)
    {
        if (string.IsNullOrWhiteSpace(html)) return string.Empty;
        var doc = new HtmlDocument();
        doc.LoadHtml(html);
        var sb = new StringBuilder();
        ConvertInlineNode(doc.DocumentNode, sb);
        return sb.ToString();
    }

    private static void ConvertInlineNode(HtmlNode node, StringBuilder sb)
    {
        foreach (var child in node.ChildNodes)
        {
            if (child.NodeType == HtmlNodeType.Text)
            {
                sb.Append(WebUtility.HtmlDecode(child.InnerText));
                continue;
            }

            if (child.NodeType != HtmlNodeType.Element) continue;

            var name = child.Name.ToLowerInvariant();
            if (name == "br")
            {
                sb.Append('\n');
                continue;
            }

            if (name == "a")
            {
                var linkText = new StringBuilder();
                ConvertInlineNode(child, linkText);
                var href = child.GetAttributeValue("href", string.Empty);
                if (!string.IsNullOrEmpty(href))
                {
                    sb.Append('[').Append(linkText).Append("](").Append(href).Append(')');
                }
                else
                {
                    sb.Append(linkText);
                }
                continue;
            }

            var style = child.GetAttributeValue("style", string.Empty).Replace(" ", string.Empty).ToLowerInvariant();
            var bold = name is "b" or "strong" || style.Contains("font-weight:bold") || style.Contains("font-weight:700");
            var italic = name is "i" or "em" || style.Contains("font-style:italic");

            var inner = new StringBuilder();
            ConvertInlineNode(child, inner);
            var content = inner.ToString();

            // Don't wrap whitespace-only runs; that produces invalid emphasis markers.
            if (!string.IsNullOrWhiteSpace(content))
            {
                if (italic) content = $"*{content}*";
                if (bold) content = $"**{content}**";
            }
            sb.Append(content);
        }
    }

    // Fetches the page-level hierarchy (notebooks > groups > sections > pages) as XML once per call.
    private XDocument LoadHierarchy()
    {
        var hierarchyXml = GetHierarchy(HierarchyScope.hsPages);
        if (string.IsNullOrEmpty(hierarchyXml))
        {
            throw new InvalidOperationException("OneNote returned an empty hierarchy. Is OneNote desktop running with notebooks loaded?");
        }

        return XDocument.Parse(hierarchyXml);
    }

    private string GetHierarchy(HierarchyScope scope)
    {
        lock (_comLock)
        {
            App.GetHierarchy(null, scope, out var xml);
            return xml ?? string.Empty;
        }
    }

    // Walks the hierarchy XML from notebook down through section groups to the target section.
    private static XElement? ResolveSectionByPath(XDocument hierarchy, string path)
    {
        var parts = path.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length < 2) return null;

        var notebook = hierarchy.Descendants(One + "Notebook")
            .FirstOrDefault(n => string.Equals((string?)n.Attribute("name"), parts[0], StringComparison.OrdinalIgnoreCase));
        if (notebook is null) return null;

        XElement current = notebook;
        for (int i = 1; i < parts.Length - 1; i++)
        {
            var next = current.Elements(One + "SectionGroup")
                .FirstOrDefault(g => string.Equals((string?)g.Attribute("name"), parts[i], StringComparison.OrdinalIgnoreCase));
            if (next is null) return null;
            current = next;
        }

        var sectionName = parts[^1];
        return current.Elements(One + "Section")
            .FirstOrDefault(s => string.Equals((string?)s.Attribute("name"), sectionName, StringComparison.OrdinalIgnoreCase));
    }

    private static DateTimeOffset ParseTime(string? value)
    {
        if (string.IsNullOrEmpty(value)) return DateTimeOffset.MinValue;
        return DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var dto)
            ? dto
            : DateTimeOffset.MinValue;
    }

    /// <summary>Decodes OneNote's HTML-encoded run text into trimmed plain text.</summary>
    public static string HtmlToPlainText(string html)
    {
        if (string.IsNullOrWhiteSpace(html)) return string.Empty;
        var doc = new HtmlDocument();
        doc.LoadHtml(html);
        var text = WebUtility.HtmlDecode(doc.DocumentNode.InnerText);
        var lines = text.Split('\n').Select(l => l.Trim()).Where(l => l.Length > 0);
        return string.Join("\n", lines);
    }
}

/// <summary>A page summary plus the category and section it was found in.</summary>
public sealed record OneNotePageInfo(
    string Id,
    string Title,
    DateTimeOffset LastModifiedUtc,
    DateTimeOffset CreatedUtc,
    string Category,
    string SectionDisplayName,
    string SectionPath);
