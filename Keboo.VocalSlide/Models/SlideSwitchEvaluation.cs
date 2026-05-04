namespace Keboo.VocalSlide.Models;

public sealed record SlideSwitchEvaluation(
    bool ShouldAdvance,
    int? TargetSlideNumber,
    double Confidence,
    string Reason);
