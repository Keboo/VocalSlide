using Keboo.VocalSlide.Models;

namespace Keboo.VocalSlide.Services;

public static class SlideAutomationIndex
{
    public static IReadOnlyList<PowerPointSlideInfo> Build(IReadOnlyList<PowerPointSlideInfo> slides) =>
        slides
            .Where(slide => !string.IsNullOrWhiteSpace(slide.AutomationPrompt))
            .OrderBy(slide => slide.SlideNumber)
            .ToArray();

    public static IReadOnlyList<PowerPointSlideInfo> BuildCandidates(IReadOnlyList<PowerPointSlideInfo> indexedSlides, int currentSlideNumber) =>
        indexedSlides
            .Where(slide => slide.SlideNumber != currentSlideNumber)
            .ToArray();

    public static PowerPointSlideInfo? FindNearest(IReadOnlyList<PowerPointSlideInfo> indexedSlides, int activeSlideNumber)
    {
        if (indexedSlides.Count == 0)
        {
            return null;
        }

        if (activeSlideNumber <= 0)
        {
            return indexedSlides[0];
        }

        return indexedSlides
            .Where(slide => slide.SlideNumber != activeSlideNumber)
            .OrderBy(slide => Math.Abs(slide.SlideNumber - activeSlideNumber))
            .ThenBy(slide => slide.SlideNumber)
            .FirstOrDefault();
    }
}
