namespace OneNoteMcp.Configuration;

/// <summary>
/// A named grouping of OneNote sections. Prompts select a category by <see cref="Name"/>
/// and the server reads pages across every section the category contains.
/// </summary>
public sealed class OneNoteCategory
{
    /// <summary>Human-friendly category name used to select it (matched case-insensitively).</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>The OneNote sections this category reads from.</summary>
    public List<OneNoteSectionRef> Sections { get; set; } = new();
}
