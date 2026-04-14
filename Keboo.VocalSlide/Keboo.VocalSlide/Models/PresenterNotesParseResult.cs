namespace Keboo.VocalSlide.Models;

public sealed record PresenterNotesParseResult(
    string PresenterNotes,
    string AutomationPrompt,
    bool HasDelimiter);
