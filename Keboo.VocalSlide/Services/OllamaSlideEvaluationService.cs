using Keboo.VocalSlide.Models;
using Microsoft.Extensions.AI;
using OllamaSharp;
using System.Text;
using System.Text.Json;

namespace Keboo.VocalSlide.Services;

public sealed class OllamaSlideEvaluationService : ILocalSlideEvaluationService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    public IList<AITool> Tools { get; } = [];

    public async Task<SlideSwitchEvaluation> EvaluateAsync(
        SlideEvaluationContext context,
        LlmOptions options,
        CancellationToken cancellationToken)
    {
        if (context.CandidateSlides.Count == 0)
        {
            return new SlideSwitchEvaluation(false, null, 0.0, "No indexed slides were available to evaluate.");
        }

        try
        {
            IChatClient client = new OllamaApiClient(new Uri(options.OllamaEndpoint), options.ModelName);

            if (Tools.Count > 0)
            {
                client = new ChatClientBuilder(client)
                    .UseFunctionInvocation()
                    .Build();
            }

            string systemPrompt = BuildSystemPrompt();
            string userPrompt = BuildUserPrompt(context);

            ChatOptions chatOptions = new()
            {
                MaxOutputTokens = options.MaxTokens,
                Temperature = 0.0f,
                Tools = Tools.Count > 0 ? Tools : null,
            };

            List<ChatMessage> messages =
            [
                new(ChatRole.System, systemPrompt),
                new(ChatRole.User, userPrompt),
            ];

            ChatResponse response = await client.GetResponseAsync(messages, chatOptions, cancellationToken)
                .ConfigureAwait(false);

            string rawResponse = response.Text ?? string.Empty;
            return ParseResponse(rawResponse, context.CandidateSlides);
        }
        catch (Exception ex)
        {
            return new SlideSwitchEvaluation(false, null, 0.0, $"Ollama evaluation failed: {ex.Message}");
        }
    }

    private static string BuildSystemPrompt()
    {
        return """
            You are deciding whether a speaker has reached the content for a PowerPoint slide in the current deck.

            Return JSON only, with no markdown fences and no extra commentary:
            {"shouldAdvance":true|false,"targetSlideNumber":<integer or null>,"confidence":<number between 0 and 1>,"reason":"<short explanation>"}

            Rules:
            - Only choose from the listed indexed slides.
            - You may choose a slide before or after the current slide if it is the best match.
            - Prefer false if the transcript is ambiguous.
            - Do not invent slide numbers or extra fields.
            - Slide numbers represent the stable deck order and must be used exactly as listed.
            - Treat the prompt text as the intent for when the speaker is ready to show the slide.
            """;
    }

    private static string BuildUserPrompt(SlideEvaluationContext context)
    {
        StringBuilder candidateBuilder = new();
        foreach (PowerPointSlideInfo candidate in context.CandidateSlides)
        {
            candidateBuilder.AppendLine($"- SlideNumber: {candidate.SlideNumber}");
            candidateBuilder.AppendLine($"  Title: {candidate.Title}");
            candidateBuilder.AppendLine($"  Prompt: {candidate.AutomationPrompt}");
        }

        return $$"""
            Current slide:
            - SlideNumber: {{context.CurrentSlide.SlideNumber}}
            - Title: {{context.CurrentSlide.Title}}
            - Prompt: {{context.CurrentSlide.AutomationPrompt}}

            Indexed slides:
            {{candidateBuilder.ToString().TrimEnd()}}

            Transcript window (spoken since the last slide change):
            {{context.TranscriptWindow}}
            """;
    }

    private static SlideSwitchEvaluation ParseResponse(string rawResponse, IReadOnlyList<PowerPointSlideInfo> candidateSlides)
    {
        if (string.IsNullOrWhiteSpace(rawResponse))
        {
            return new SlideSwitchEvaluation(false, null, 0.0, "Ollama produced an empty response.");
        }

        string? json = ExtractJson(rawResponse);
        if (string.IsNullOrWhiteSpace(json))
        {
            return new SlideSwitchEvaluation(false, null, 0.0, $"Ollama returned a non-JSON response: {rawResponse.Trim()}");
        }

        LlmDecisionResponse? response = JsonSerializer.Deserialize<LlmDecisionResponse>(json, JsonOptions);
        if (response is null)
        {
            return new SlideSwitchEvaluation(false, null, 0.0, "The Ollama response could not be parsed.");
        }

        int? targetSlide = response.ShouldAdvance ? response.TargetSlideNumber : null;
        if (targetSlide is not null && candidateSlides.All(slide => slide.SlideNumber != targetSlide.Value))
        {
            return new SlideSwitchEvaluation(false, null, 0.0, "Ollama selected a slide outside the indexed slide set.");
        }

        double confidence = Math.Clamp(response.Confidence, 0.0, 1.0);
        string reason = string.IsNullOrWhiteSpace(response.Reason)
            ? "No explanation was provided."
            : response.Reason.Trim();

        return new SlideSwitchEvaluation(response.ShouldAdvance, targetSlide, confidence, reason);
    }

    private static string? ExtractJson(string rawResponse)
    {
        int start = rawResponse.IndexOf('{');
        int end = rawResponse.LastIndexOf('}');

        if (start < 0 || end < start)
        {
            return null;
        }

        return rawResponse[start..(end + 1)];
    }

    private sealed record LlmDecisionResponse(
        bool ShouldAdvance,
        int? TargetSlideNumber,
        double Confidence,
        string? Reason);
}
