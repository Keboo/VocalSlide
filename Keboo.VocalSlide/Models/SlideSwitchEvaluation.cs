namespace Keboo.VocalSlide.Models;

public sealed record SlideSwitchEvaluation(
    int? TargetSlideNumber,
    int Confidence,
    string Reason)

{
    public static SlideSwitchEvaluation Fail(string reason) => new(null, 0, reason);
}
