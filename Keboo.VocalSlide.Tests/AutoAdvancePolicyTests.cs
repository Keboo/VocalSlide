using Keboo.VocalSlide.Models;
using Keboo.VocalSlide.Services;

namespace Keboo.VocalSlide.Tests;

public class AutoAdvancePolicyTests
{
    [Test]
    public async Task Evaluate_WhenConfidenceIsBelowThreshold_DoesNotAdvance()
    {
        AutoAdvancePolicy policy = new();
        SlideSwitchEvaluation evaluation = new(true, 4, 0.42, "The transcript only weakly matches the next slide.");

        AutoAdvanceDecisionOutcome result = policy.Evaluate(
            evaluation,
            new HashSet<int> { 4, 5 },
            0.75,
            DateTimeOffset.UtcNow,
            null,
            TimeSpan.FromSeconds(4));

        await Assert.That(result.ShouldAdvance).IsFalse();
        await Assert.That(result.TargetSlideNumber.HasValue).IsFalse();
    }

    [Test]
    public async Task Evaluate_WhenDecisionIsCoolingDown_DoesNotAdvance()
    {
        AutoAdvancePolicy policy = new();
        SlideSwitchEvaluation evaluation = new(true, 3, 0.98, "The speaker is clearly on the next topic.");

        AutoAdvanceDecisionOutcome result = policy.Evaluate(
            evaluation,
            new HashSet<int> { 3, 4 },
            0.75,
            new DateTimeOffset(2026, 1, 1, 12, 0, 2, TimeSpan.Zero),
            new DateTimeOffset(2026, 1, 1, 12, 0, 0, TimeSpan.Zero),
            TimeSpan.FromSeconds(5));

        await Assert.That(result.ShouldAdvance).IsFalse();
        await Assert.That(result.TargetSlideNumber.HasValue).IsFalse();
    }

    [Test]
    public async Task Evaluate_WhenTargetSlideIsOutsideCandidateWindow_DoesNotAdvance()
    {
        AutoAdvancePolicy policy = new();
        SlideSwitchEvaluation evaluation = new(true, 8, 0.99, "The evaluator jumped too far ahead.");

        AutoAdvanceDecisionOutcome result = policy.Evaluate(
            evaluation,
            new HashSet<int> { 3, 4 },
            0.75,
            DateTimeOffset.UtcNow,
            null,
            TimeSpan.FromSeconds(5));

        await Assert.That(result.ShouldAdvance).IsFalse();
        await Assert.That(result.TargetSlideNumber.HasValue).IsFalse();
    }

    [Test]
    public async Task Evaluate_WhenDecisionIsConfidentAndAllowed_AdvancesTheTargetSlide()
    {
        AutoAdvancePolicy policy = new();
        SlideSwitchEvaluation evaluation = new(true, 5, 0.94, "The speaker has started discussing slide five.");

        AutoAdvanceDecisionOutcome result = policy.Evaluate(
            evaluation,
            new HashSet<int> { 5, 6 },
            0.75,
            DateTimeOffset.UtcNow,
            null,
            TimeSpan.FromSeconds(4));

        await Assert.That(result.ShouldAdvance).IsTrue();
        await Assert.That(result.TargetSlideNumber).IsEqualTo(5);
        await Assert.That(result.Confidence).IsEqualTo(0.94);
    }

    [Test]
    public async Task Evaluate_WhenDecisionTargetsEarlierIndexedSlide_AllowsBackwardJump()
    {
        AutoAdvancePolicy policy = new();
        SlideSwitchEvaluation evaluation = new(true, 2, 0.91, "The speaker returned to an earlier part of the deck.");

        AutoAdvanceDecisionOutcome result = policy.Evaluate(
            evaluation,
            new HashSet<int> { 2, 4, 6 },
            0.75,
            DateTimeOffset.UtcNow,
            null,
            TimeSpan.FromSeconds(4));

        await Assert.That(result.ShouldAdvance).IsTrue();
        await Assert.That(result.TargetSlideNumber).IsEqualTo(2);
        await Assert.That(result.Confidence).IsEqualTo(0.91);
    }
}
