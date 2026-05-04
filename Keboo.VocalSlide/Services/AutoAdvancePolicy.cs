using Keboo.VocalSlide.Models;

namespace Keboo.VocalSlide.Services;

public sealed class AutoAdvancePolicy
{
    public AutoAdvanceDecisionOutcome Evaluate(
        SlideSwitchEvaluation evaluation,
        IReadOnlySet<int> candidateSlideNumbers,
        double confidenceThreshold,
        DateTimeOffset evaluationTime,
        DateTimeOffset? lastAdvanceAt,
        TimeSpan cooldown)
    {
        if (!evaluation.ShouldAdvance)
        {
            return AutoAdvanceDecisionOutcome.Skip($"Held current slide. {evaluation.Reason}");
        }

        if (evaluation.TargetSlideNumber is null)
        {
            return AutoAdvanceDecisionOutcome.Skip("Held current slide. The evaluator requested a slide change without a target slide.");
        }

        if (!candidateSlideNumbers.Contains(evaluation.TargetSlideNumber.Value))
        {
            return AutoAdvanceDecisionOutcome.Skip("Held current slide. The selected target was outside the indexed slide set.");
        }

        if (evaluation.Confidence < confidenceThreshold)
        {
            return AutoAdvanceDecisionOutcome.Skip(
                $"Held current slide. Confidence {evaluation.Confidence:0.00} is below the threshold {confidenceThreshold:0.00}.");
        }

        if (lastAdvanceAt is not null && evaluationTime - lastAdvanceAt.Value < cooldown)
        {
            return AutoAdvanceDecisionOutcome.Skip("Held current slide. Auto-advance is still cooling down.");
        }

        return AutoAdvanceDecisionOutcome.Advance(
            evaluation.TargetSlideNumber.Value,
            evaluation.Confidence,
            $"Move to slide {evaluation.TargetSlideNumber.Value} at confidence {evaluation.Confidence:0.00}. {evaluation.Reason}");
    }
}
