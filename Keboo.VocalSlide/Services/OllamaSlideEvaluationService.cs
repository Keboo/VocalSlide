using System.Text;

using Keboo.VocalSlide.Models;

using Microsoft.Extensions.AI;

using OllamaSharp;

namespace Keboo.VocalSlide.Services;

public sealed class OllamaSlideEvaluationService : ILocalSlideEvaluationService
{
    private static readonly TimeSpan RecentTranscriptWindow = TimeSpan.FromSeconds(60);

    public IList<AITool> Tools { get; } = [];

    public async Task<SlideSwitchEvaluation> EvaluateAsync(
        SlideEvaluationContext context,
        LlmOptions options,
        CancellationToken cancellationToken)
    {
        if (context.CandidateSlides.Count == 0)
        {
            return SlideSwitchEvaluation.Fail("No indexed slides were available to evaluate.");
        }

        try
        {
            IChatClient client = new OllamaApiClient(new Uri(options.OllamaEndpoint), options.ModelName);

            List<AITool> tools = [.. Tools];
            tools.Add(TranscriptSearchTool.Create(context.FullTranscript));

            if (tools.Count > 0)
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
                Temperature = 0.2f,
                ResponseFormat = ChatResponseFormat.ForJsonSchema<DecisionResponse>(),
                Tools = tools.Count > 0 ? tools : null,
            };

            List<ChatMessage> messages =
            [
                new(ChatRole.System, systemPrompt),
                new(ChatRole.User, userPrompt),
            ];

            ChatResponse<DecisionResponse> response = await client.GetResponseAsync<DecisionResponse>(messages, chatOptions, cancellationToken:cancellationToken)
                .ConfigureAwait(false);

            return ParseResponse(response, context.CandidateSlides);
        }
        catch (Exception ex)
        {
            return SlideSwitchEvaluation.Fail($"Ollama evaluation failed: {ex.Message}");
        }
    }

    private static string BuildSystemPrompt()
    {
        return $$"""
            You are deciding which slide to show during a PowerPoint presentation.

            Rules:
            - The target slide number must come from the indexed slide numbers.
            - You must provide a reason for picking a slide.
            - Provide a confidence value for picking the slide from 0 to 100, where 0 is no confidence and 100 is absolute confidence.
            - Give the highest priority to the next slide in the deck.
            - Avoid moving backwards through the deck unless the transcript clearly indicates a return to a previous slide.
            - If the user explicitly requests to move to the next or previous slide, you must do so.
            
            - You have access to a search_transcript tool that searches the full transcript using regex.
            - If you need more context to make a decision, use search_transcript to query earlier parts of the transcript.
            """;
    }

    private static string BuildUserPrompt(SlideEvaluationContext context)
    {
        StringBuilder candidateBuilder = new();
        foreach (PowerPointSlideInfo candidate in context.CandidateSlides)
        {
            candidateBuilder.AppendLine($"- SlideNumber: {candidate.SlideNumber}");
            candidateBuilder.AppendLine($"  Prompt: {candidate.AutomationPrompt}");
        }

        string recentTranscript = BuildRecentTranscript(context.FullTranscript);

        return $$"""
            Current slide:
            - SlideNumber: {{context.CurrentSlide.SlideNumber}}
            - Prompt: {{context.CurrentSlide.AutomationPrompt}}

            Indexed slides:
            {{candidateBuilder.ToString().TrimEnd()}}

            Transcript (last 60 seconds):
            {{recentTranscript}}
            """;
    }

    private static string BuildRecentTranscript(IReadOnlyList<TranscriptEntry> fullTranscript)
    {
        if (fullTranscript.Count == 0)
        {
            return string.Empty;
        }

        DateTimeOffset cutoff = DateTimeOffset.Now - RecentTranscriptWindow;
        var recentEntries = new List<TranscriptEntry>();

        for (int i = fullTranscript.Count - 1; i >= 0; i--)
        {
            TranscriptEntry entry = fullTranscript[i];
            if (entry.Timestamp >= cutoff)
            {
                recentEntries.Add(entry);
            }
            else
            {
                break;
            }
        }

        if (recentEntries.Count == 0)
        {
            return fullTranscript[^1].Text;
        }

        StringBuilder sb = new();
        for (int i = recentEntries.Count - 1; i >= 0; i--)
        {
            sb.AppendLine($"[{recentEntries[i].Timestamp:HH:mm:ss}] {recentEntries[i].Text}");
        }

        return sb.ToString().TrimEnd();
    }

    private static SlideSwitchEvaluation ParseResponse(ChatResponse<DecisionResponse> response, IReadOnlyList<PowerPointSlideInfo> candidateSlides)
    {
        if (response is null || response.Result is null)
        {
            return SlideSwitchEvaluation.Fail("Ollama produced an empty response.");
        }

        DecisionResponse? result = response.Result;
        if (result is null)
        {
            return SlideSwitchEvaluation.Fail("Ollama returned a non-JSON response.");
        }

        int? targetSlide = result.TargetSlideNumber;
        if (targetSlide is not null && candidateSlides.All(slide => slide.SlideNumber != targetSlide.Value))
        {
            return SlideSwitchEvaluation.Fail("Ollama selected a slide outside the indexed slide set.");
        }

        int confidence = Math.Clamp(result.Confidence, 0, 100);
        string reason = result.Reason?.Trim() ?? "";

        return new SlideSwitchEvaluation(targetSlide, confidence, reason);
    }

    private sealed record DecisionResponse(
        int? TargetSlideNumber,
        int Confidence,
        string? Reason);
}
