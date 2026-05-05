namespace Keboo.VocalSlide.Models;

public sealed record AutoAdvanceDecisionOutcome(
    int? TargetSlideNumber,
    int Confidence,
    string Summary)
{
    public static AutoAdvanceDecisionOutcome Skip(string summary) =>
        new(null, 0, summary);

    public static AutoAdvanceDecisionOutcome Advance(int slideNumber, int confidence, string summary) =>
        new(slideNumber, confidence, summary);
}
