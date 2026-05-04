using Keboo.VocalSlide.Models;
using LLama;
using LLama.Common;
using LLama.Sampling;
using System.IO;
using System.Text;
using System.Text.Json;

namespace Keboo.VocalSlide.Services;

public sealed class LlamaSlideEvaluationService : ILocalSlideEvaluationService, IDisposable
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    private static readonly Grammar JsonDecisionGrammar = new(
        """
        root ::= "{" ws "\"shouldAdvance\"" ":" ws boolean "," ws "\"targetSlideNumber\"" ":" ws nullable-integer "," ws "\"confidence\"" ":" ws number "," ws "\"reason\"" ":" ws string "}" ws
        boolean ::= "true" | "false"
        nullable-integer ::= integer | "null" ws
        integer ::= ("-"? ([0-9] | [1-9] [0-9]{0,15})) ws
        string ::= "\"" (
          [^"\\\x7F\x00-\x1F] |
          "\\" (["\\bfnrt] | "u" [0-9a-fA-F]{4})
        )* "\"" ws
        number ::= ("-"? ([0-9] | [1-9] [0-9]{0,15})) ("." [0-9]+)? ([eE] [-+]? [0-9] [1-9]{0,15})? ws
        ws ::= | " " | "\n" [ \t]{0,20}
        """,
        "root");

    private readonly SemaphoreSlim _modelLock = new(1, 1);

    private LLamaWeights? _model;
    private StatelessExecutor? _executor;
    private string? _activeModelPath;
    private int _activeContextSize;
    private int _activeGpuLayerCount;

    public async Task<SlideSwitchEvaluation> EvaluateAsync(
        SlideEvaluationContext context,
        LlmOptions options,
        CancellationToken cancellationToken)
    {
        if (context.CandidateSlides.Count == 0)
        {
            return new SlideSwitchEvaluation(false, null, 0.0, "No indexed slides were available to evaluate.");
        }

        await _modelLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            EnsureExecutor(options);

            string prompt = BuildPrompt(context);
            InferenceParams inferenceParams = new()
            {
                MaxTokens = options.MaxTokens,
                SamplingPipeline = new DefaultSamplingPipeline
                {
                    Temperature = 0.0f,
                    Grammar = JsonDecisionGrammar,
                    GrammarOptimization = DefaultSamplingPipeline.GrammarOptimizationMode.Basic,
                },
            };

            StringBuilder responseBuilder = new();
            await foreach (string token in _executor!.InferAsync(prompt, inferenceParams).ConfigureAwait(false))
            {
                responseBuilder.Append(token);
            }

            return ParseResponse(responseBuilder.ToString(), context.CandidateSlides);
        }
        catch (FileNotFoundException)
        {
            throw;
        }
        catch (Exception ex)
        {
            return new SlideSwitchEvaluation(false, null, 0.0, $"The local LLM could not evaluate the transcript: {ex.Message}");
        }
        finally
        {
            _modelLock.Release();
        }
    }

    private void EnsureExecutor(LlmOptions options)
    {
        if (_executor is not null &&
            string.Equals(_activeModelPath, options.ModelPath, StringComparison.OrdinalIgnoreCase) &&
            _activeContextSize == options.ContextSize &&
            _activeGpuLayerCount == options.GpuLayerCount)
        {
            return;
        }

        DisposeModel();

        if (!File.Exists(options.ModelPath))
        {
            throw new FileNotFoundException("LLM model file not found.", options.ModelPath);
        }

        ModelParams modelParams = new(options.ModelPath)
        {
            ContextSize = (uint)options.ContextSize,
            GpuLayerCount = options.GpuLayerCount,
        };

        _model = LLamaWeights.LoadFromFile(modelParams);
        _executor = new StatelessExecutor(_model, modelParams);
        _activeModelPath = options.ModelPath;
        _activeContextSize = options.ContextSize;
        _activeGpuLayerCount = options.GpuLayerCount;
    }

    private static string BuildPrompt(SlideEvaluationContext context)
    {
        StringBuilder candidateBuilder = new();
        foreach (PowerPointSlideInfo candidate in context.CandidateSlides)
        {
            candidateBuilder.AppendLine($"- SlideNumber: {candidate.SlideNumber}");
            candidateBuilder.AppendLine($"  Title: {candidate.Title}");
            candidateBuilder.AppendLine($"  Prompt: {candidate.AutomationPrompt}");
        }

        return $$"""
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
            return new SlideSwitchEvaluation(false, null, 0.0, "The local LLM produced an empty response.");
        }

        string? json = ExtractJson(rawResponse);
        if (string.IsNullOrWhiteSpace(json))
        {
            return new SlideSwitchEvaluation(false, null, 0.0, $"The local LLM returned a non-JSON response: {rawResponse.Trim()}");
        }

        LlmDecisionResponse? response = JsonSerializer.Deserialize<LlmDecisionResponse>(json, JsonOptions);
        if (response is null)
        {
            return new SlideSwitchEvaluation(false, null, 0.0, "The local LLM response could not be parsed.");
        }

        int? targetSlide = response.ShouldAdvance ? response.TargetSlideNumber : null;
        if (targetSlide is not null && candidateSlides.All(slide => slide.SlideNumber != targetSlide.Value))
        {
            return new SlideSwitchEvaluation(false, null, 0.0, "The local LLM selected a slide outside the indexed slide set.");
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

    private void DisposeModel()
    {
        _executor = null;
        _model?.Dispose();
        _model = null;
        _activeModelPath = null;
        _activeContextSize = 0;
        _activeGpuLayerCount = 0;
    }

    public void Dispose()
    {
        DisposeModel();
        _modelLock.Dispose();
    }

    private sealed record LlmDecisionResponse(
        bool ShouldAdvance,
        int? TargetSlideNumber,
        double Confidence,
        string? Reason);
}
