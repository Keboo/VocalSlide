namespace Keboo.VocalSlide.Models;

public sealed record SlideEvaluationContext(
    PowerPointSlideInfo CurrentSlide,
    IReadOnlyList<PowerPointSlideInfo> CandidateSlides,
    string TranscriptWindow,
    IReadOnlyList<TranscriptEntry> FullTranscript);
