using Keboo.VocalSlide.Models;

namespace Keboo.VocalSlide.Services;

public sealed class AutoAdvancePolicy
{
    public AutoAdvanceDecisionOutcome Evaluate(
        SlideSwitchEvaluation evaluation,
        IReadOnlySet<int> candidateSlideNumbers,
        int confidenceThreshold,
        DateTimeOffset evaluationTime,
        DateTimeOffset? lastAdvanceAt,
        TimeSpan cooldown)
    {
        if (evaluation.TargetSlideNumber is null)
        {
            return AutoAdvanceDecisionOutcome.Skip("Held current slide. No target slide indicated.");
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

        //if (string.IsNullOrWhiteSpace(evaluation.Reason))
        //{
        //    return AutoAdvanceDecisionOutcome.Skip(
        //        $"Held current slide. Confidence {evaluation.Confidence:0.00} is above the threshold but no reason to switch was provided.");
        //}

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
