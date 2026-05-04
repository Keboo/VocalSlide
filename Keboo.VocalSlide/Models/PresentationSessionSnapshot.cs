namespace Keboo.VocalSlide.Models;

public sealed record PresentationSessionSnapshot(
    string PresentationName,
    SlideShowState SlideShowState,
    IReadOnlyList<PowerPointSlideInfo> Slides);
