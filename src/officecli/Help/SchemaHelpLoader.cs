// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;

namespace OfficeCli.Help;

/// <summary>
/// Locates and loads help schemas from the schemas/help tree. Resolves format
/// aliases (word/excel/ppt) and element aliases declared inside each schema.
/// </summary>
internal static class SchemaHelpLoader
{
    private static readonly string[] CanonicalFormats = { "docx", "xlsx", "pptx" };

    private static readonly Dictionary<string, string> FormatAliases =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["docx"] = "docx",
            ["word"] = "docx",
            ["xlsx"] = "xlsx",
            ["excel"] = "xlsx",
            ["pptx"] = "pptx",
            ["ppt"] = "pptx",
            ["powerpoint"] = "pptx",
        };

    private static string? _cachedRoot;

    internal static string LocateSchemasRoot()
    {
        if (_cachedRoot != null) return _cachedRoot;

        // 1. AppContext.BaseDirectory direct: schemas ship as Content next to
        //    the built binary (bin/Debug/.../ or published single-file location).
        var baseDir = AppContext.BaseDirectory;
        var direct = Path.Combine(baseDir, "schemas", "help");
        if (Directory.Exists(direct))
        {
            _cachedRoot = direct;
            return direct;
        }

        // 2. Walk up from AppContext.BaseDirectory looking for schemas/help
        //    (same logic as SchemaContractTests). Handles dev-tree `dotnet run`
        //    where bin/ is several levels below the repo root.
        var dir = baseDir;
        for (int i = 0; i < 10 && dir is not null; i++)
        {
            var candidate = Path.Combine(dir, "schemas", "help");
            if (Directory.Exists(candidate))
            {
                _cachedRoot = candidate;
                return candidate;
            }
            dir = Path.GetDirectoryName(dir);
        }

        throw new DirectoryNotFoundException(
            "Could not locate schemas/help/ starting from " + baseDir);
    }

    internal static IReadOnlyList<string> ListFormats() => CanonicalFormats;

    /// <summary>
    /// Normalize a user-supplied format token to canonical docx/xlsx/pptx.
    /// Throws InvalidOperationException with a suggestion if unknown.
    /// </summary>
    internal static string NormalizeFormat(string input)
    {
        if (FormatAliases.TryGetValue(input, out var canonical)) return canonical;

        // Suggest closest format alias
        var best = ClosestMatch(input, FormatAliases.Keys);
        var suggestion = best != null ? $" Did you mean: {best}?" : "";
        throw new InvalidOperationException(
            $"error: unknown format '{input}'.{suggestion}\n" +
            $"Use: officecli help");
    }

    internal static IReadOnlyList<string> ListElements(string format)
    {
        var canonical = NormalizeFormat(format);
        var dir = Path.Combine(LocateSchemasRoot(), canonical);
        if (!Directory.Exists(dir))
            throw new DirectoryNotFoundException($"Schema directory missing: {dir}");

        return Directory.GetFiles(dir, "*.json")
            .Select(Path.GetFileNameWithoutExtension)
            .Where(n => !string.IsNullOrEmpty(n))
            .Select(n => n!)
            .OrderBy(n => n, StringComparer.Ordinal)
            .ToList();
    }

    /// <summary>
    /// Load a schema for (format, element). Element can be the filename stem
    /// or any alias declared in another schema's "aliases" entry (rare, mostly
    /// a property-level concept, but checked for completeness).
    /// </summary>
    internal static JsonDocument LoadSchema(string format, string element)
    {
        var canonical = NormalizeFormat(format);
        var dir = Path.Combine(LocateSchemasRoot(), canonical);
        var elements = ListElements(canonical);

        // 1. Exact filename match (case-insensitive).
        var match = elements.FirstOrDefault(
            e => string.Equals(e, element, StringComparison.OrdinalIgnoreCase));
        if (match != null)
        {
            var full = Path.Combine(dir, match + ".json");
            return JsonDocument.Parse(File.ReadAllText(full));
        }

        // 2. Unknown element — suggest closest match.
        var best = ClosestMatch(element, elements);
        var suggestion = best != null ? $"\nDid you mean: {best}?" : "";
        throw new InvalidOperationException(
            $"error: unknown element '{element}' for format '{canonical}'.{suggestion}\n" +
            $"Use: officecli help {canonical}");
    }

    /// <summary>
    /// Check whether a schema's top-level operations[verb] is true. Used by
    /// `officecli help &lt;format&gt; &lt;verb&gt;` to filter the element list.
    /// </summary>
    internal static bool ElementSupportsVerb(string format, string element, string verb)
    {
        try
        {
            using var doc = LoadSchema(format, element);
            if (doc.RootElement.TryGetProperty("operations", out var ops)
                && ops.TryGetProperty(verb, out var v)
                && v.ValueKind == JsonValueKind.True)
            {
                return true;
            }
        }
        catch
        {
            // Swallow — a bad schema shouldn't kill the filter.
        }
        return false;
    }

    /// <summary>
    /// Suggest the closest candidate from <paramref name="candidates"/> to
    /// <paramref name="input"/> using substring + Levenshtein. Returns null
    /// if no candidate is close enough.
    /// </summary>
    private static string? ClosestMatch(string input, IEnumerable<string> candidates)
    {
        var lower = input.ToLowerInvariant();

        // Prefer substring hit (common for user typos like `paragrah`).
        var substringHit = candidates.FirstOrDefault(
            c => c.Contains(lower, StringComparison.OrdinalIgnoreCase)
                 || lower.Contains(c, StringComparison.OrdinalIgnoreCase));

        string? best = null;
        int bestDist = int.MaxValue;
        foreach (var c in candidates)
        {
            var dist = LevenshteinDistance(lower, c.ToLowerInvariant());
            // Accept distance up to max(2, len/3) — same rule CommandBuilder uses.
            var maxDist = Math.Max(2, lower.Length / 3);
            if (dist <= maxDist && dist < bestDist)
            {
                best = c;
                bestDist = dist;
            }
        }

        return best ?? substringHit;
    }

    private static int LevenshteinDistance(string s, string t)
    {
        if (s.Length == 0) return t.Length;
        if (t.Length == 0) return s.Length;

        var d = new int[s.Length + 1, t.Length + 1];
        for (int i = 0; i <= s.Length; i++) d[i, 0] = i;
        for (int j = 0; j <= t.Length; j++) d[0, j] = j;

        for (int i = 1; i <= s.Length; i++)
        {
            for (int j = 1; j <= t.Length; j++)
            {
                int cost = s[i - 1] == t[j - 1] ? 0 : 1;
                d[i, j] = Math.Min(
                    Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                    d[i - 1, j - 1] + cost);
            }
        }

        return d[s.Length, t.Length];
    }
}
