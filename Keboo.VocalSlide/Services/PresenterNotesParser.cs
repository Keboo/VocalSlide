using Keboo.VocalSlide.Models;
using System.Text.RegularExpressions;

namespace Keboo.VocalSlide.Services;

public sealed partial class PresenterNotesParser : IPresenterNotesParser
{
    [GeneratedRegex(@"^\s*-{3,}\s*$", RegexOptions.Multiline)]
    private static partial Regex DelimiterRegex();

    public PresenterNotesParseResult Parse(string? rawNotes)
    {
        if (string.IsNullOrWhiteSpace(rawNotes))
        {
            return new PresenterNotesParseResult(string.Empty, string.Empty, false);
        }

        string normalized = NormalizeLineEndings(rawNotes).Trim();
        Match match = DelimiterRegex().Match(normalized);

        if (!match.Success)
        {
            return new PresenterNotesParseResult(normalized, string.Empty, false);
        }

        string presenterNotes = normalized[..match.Index].Trim();
        string automationPrompt = normalized[(match.Index + match.Length)..].Trim();

        return new PresenterNotesParseResult(presenterNotes, automationPrompt, true);
    }

    private static string NormalizeLineEndings(string rawNotes) =>
        rawNotes
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Replace("\u2028", "\n", StringComparison.Ordinal)
            .Replace("\u2029", "\n", StringComparison.Ordinal);
}
