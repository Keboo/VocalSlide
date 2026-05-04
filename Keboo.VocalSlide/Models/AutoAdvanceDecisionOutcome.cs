namespace Keboo.VocalSlide.Models;

public sealed record AutoAdvanceDecisionOutcome(
    bool ShouldAdvance,
    int? TargetSlideNumber,
    double Confidence,
    string Summary)
{
    public static AutoAdvanceDecisionOutcome Skip(string summary) =>
        new(false, null, 0.0, summary);

    public static AutoAdvanceDecisionOutcome Advance(int slideNumber, double confidence, string summary) =>
        new(true, slideNumber, confidence, summary);
}
