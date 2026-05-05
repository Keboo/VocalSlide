using Keboo.VocalSlide.Models;
using Microsoft.Extensions.AI;
using System.ComponentModel;
using System.Text;
using System.Text.RegularExpressions;

namespace Keboo.VocalSlide.Services;

public static class TranscriptSearchTool
{
    public static AITool Create(IReadOnlyList<TranscriptEntry> fullTranscript)
    {
        return AIFunctionFactory.Create(
            [Description("Search the full transcript using a .NET regular expression pattern. Returns matching lines with timestamps. Use this to find earlier context not included in the recent transcript window.")]
            (
                [Description("A .NET regular expression pattern to search for in the transcript text. Case-insensitive.")] string pattern
            ) =>
            {
                try
                {
                    Regex regex = new(pattern, RegexOptions.IgnoreCase, TimeSpan.FromSeconds(2));
                    StringBuilder results = new();
                    int matchCount = 0;

                    foreach (TranscriptEntry entry in fullTranscript)
                    {
                        if (regex.IsMatch(entry.Text))
                        {
                            results.AppendLine($"[{entry.Timestamp:HH:mm:ss}] {entry.Text}");
                            matchCount++;
                        }
                    }

                    return matchCount > 0
                        ? $"Found {matchCount} matching entries:\n{results.ToString().TrimEnd()}"
                        : "No matches found.";
                }
                catch (RegexParseException ex)
                {
                    return $"Invalid regex pattern: {ex.Message}";
                }
                catch (RegexMatchTimeoutException)
                {
                    return "The regex search timed out. Try a simpler pattern.";
                }
            },
            name: "search_transcript",
            description: "Search the full transcript using a .NET regular expression pattern. Returns matching lines with timestamps. Use this to find earlier context not included in the recent transcript window.");
    }
}
