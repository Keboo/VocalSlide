namespace Keboo.VocalSlide.Models;

public sealed record PowerPointSlideInfo(
    int SlideNumber,
    string Title,
    string PresenterNotes,
    string AutomationPrompt)
{
    public string PromptPreview => string.IsNullOrWhiteSpace(AutomationPrompt)
        ? "(no automation prompt)"
        : AutomationPrompt;
}
