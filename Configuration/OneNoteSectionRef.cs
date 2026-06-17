namespace OneNoteMcp.Configuration;

/// <summary>
/// Points at a single OneNote section by its hierarchical display path.
/// </summary>
public sealed class OneNoteSectionRef
{
    /// <summary>
    /// Hierarchical path to the section: "Notebook Name/Group/SubGroup/Section Name".
    /// OneNote desktop assigns its own internal section IDs, so we identify by path instead.
    /// </summary>
    public string Path { get; set; } = string.Empty;

    /// <summary>Human-readable label for logging / tool output. Defaults to the last path segment.</summary>
    public string? DisplayName { get; set; }

    /// <summary>
    /// The label to surface to callers: <see cref="DisplayName"/> when set, otherwise the
    /// final segment of <see cref="Path"/> (the section name).
    /// </summary>
    public string EffectiveDisplayName
    {
        get
        {
            if (!string.IsNullOrWhiteSpace(DisplayName)) return DisplayName!;
            var parts = Path.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            return parts.Length > 0 ? parts[^1] : Path;
        }
    }
}
