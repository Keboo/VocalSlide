namespace Keboo.VocalSlide.Models;

public sealed record LlmOptions(
    string ModelPath,
    int ContextSize = 2048,
    int GpuLayerCount = 0,
    int MaxTokens = 160);
