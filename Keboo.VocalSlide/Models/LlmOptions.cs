namespace Keboo.VocalSlide.Models;

public sealed record LlmOptions(
    string OllamaEndpoint = "http://localhost:11434",
    string ModelName = "qwen2.5:0.5b",
    int MaxTokens = 200);
