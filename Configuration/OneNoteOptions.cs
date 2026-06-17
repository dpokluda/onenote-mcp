namespace OneNoteMcp.Configuration;

/// <summary>
/// Root configuration object bound from the "OneNote" section of <c>appsettings.json</c>.
/// </summary>
public sealed class OneNoteOptions
{
    /// <summary>Name of the configuration section these options bind from.</summary>
    public const string SectionName = "OneNote";

    /// <summary>
    /// Named categories, each grouping one or more OneNote sections. Prompts reference a
    /// category (e.g. "TSG" or "On-Call Notes") and the server reads across every section in it.
    /// </summary>
    public List<OneNoteCategory> Categories { get; set; } = new();
}
