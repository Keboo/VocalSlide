using Keboo.VocalSlide.Models;

namespace Keboo.VocalSlide.Services;

public interface ILocalSlideEvaluationService
{
    Task<SlideSwitchEvaluation> EvaluateAsync(
        SlideEvaluationContext context,
        LlmOptions options,
        CancellationToken cancellationToken);
}
