using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Keboo.VocalSlide.Models;
using Keboo.VocalSlide.Services;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows.Threading;

namespace Keboo.VocalSlide;

public partial class MainWindowViewModel : ObservableObject
{
    private const int MaxTranscriptChunks = 10;
    private const int MaxTranscriptCharacters = 4000;

    private readonly IPowerPointSessionService _powerPointSessionService;
    private readonly ILocalTranscriptionService _transcriptionService;
    private readonly ILocalSlideEvaluationService _slideEvaluationService;
    private readonly IModelDownloadService _modelDownloadService;
    private readonly AutoAdvancePolicy _autoAdvancePolicy;
    private readonly Dispatcher _dispatcher;
    private readonly object _transcriptLock = new();
    private readonly Queue<string> _transcriptChunks = new();
    private readonly SemaphoreSlim _evaluationGate = new(1, 1);

    private IReadOnlyList<PowerPointSlideInfo> _slideCache = Array.Empty<PowerPointSlideInfo>();
    private IReadOnlyList<PowerPointSlideInfo> _automationPromptIndex = Array.Empty<PowerPointSlideInfo>();
    private CancellationTokenSource? _automationCts;
    private DateTimeOffset? _lastAdvanceAt;

    public MainWindowViewModel(
        IPowerPointSessionService powerPointSessionService,
        ILocalTranscriptionService transcriptionService,
        ILocalSlideEvaluationService slideEvaluationService,
        IModelDownloadService modelDownloadService,
        AutoAdvancePolicy autoAdvancePolicy,
        Dispatcher dispatcher)
    {
        _powerPointSessionService = powerPointSessionService;
        _transcriptionService = transcriptionService;
        _slideEvaluationService = slideEvaluationService;
        _modelDownloadService = modelDownloadService;
        _autoAdvancePolicy = autoAdvancePolicy;
        _dispatcher = dispatcher;

        ModelStorageDirectory = _modelDownloadService.StorageDirectory;

        foreach (DownloadableModelOption option in ModelCatalog.WhisperModels)
        {
            WhisperDownloadOptions.Add(option);
        }

        SelectedWhisperDownloadOption = ModelCatalog.DefaultWhisper;
        WhisperModelPath = _modelDownloadService.GetStoredFilePath(ModelCatalog.DefaultWhisper.FileName);

        LlmDownloadUrl = ModelCatalog.DefaultLlm.DownloadUrl;
        LlmModelPath = _modelDownloadService.GetStoredFilePath(ModelCatalog.DefaultLlm.FileName);

        RefreshModelAvailability();
    }

    public ObservableCollection<PowerPointSlideInfo> Slides { get; } = new();

    public ObservableCollection<DownloadableModelOption> WhisperDownloadOptions { get; } = new();

    [ObservableProperty]
    private string _modelStorageDirectory = string.Empty;

    [ObservableProperty]
    private string _presentationName = "Not connected";

    [ObservableProperty]
    private string _connectionSummary = "Waiting for a running PowerPoint slide show.";

    [ObservableProperty]
    private string _slideShowStateText = "No slide show detected.";

    [ObservableProperty]
    private string _statusMessage = "Open PowerPoint, start the slide show, then refresh the deck. Models can be downloaded into the local app-data folder.";

    [ObservableProperty]
    private string _currentSlideTitle = "No active slide";

    [ObservableProperty]
    private string _currentSlidePrompt = "Refresh an active presentation to load the presenter notes prompts.";

    [ObservableProperty]
    private string _nextSlideTitle = "No indexed prompt slide";

    [ObservableProperty]
    private string _nextSlidePrompt = "No prompt available.";

    [ObservableProperty]
    private string _transcriptWindow = string.Empty;

    [ObservableProperty]
    private string _lastDecision = "No decisions yet.";

    [ObservableProperty]
    private DownloadableModelOption? _selectedWhisperDownloadOption;

    [ObservableProperty]
    private string _whisperDownloadStatus = "Select or download a Whisper model.";

    [ObservableProperty]
    private double _whisperDownloadPercent;

    [ObservableProperty]
    private bool _isWhisperDownloadInProgress;

    [ObservableProperty]
    private string _whisperModelPath = string.Empty;

    [ObservableProperty]
    private string _llmDownloadUrl = string.Empty;

    [ObservableProperty]
    private string _llmDownloadStatus = "Paste a direct GGUF URL or use the recommended default.";

    [ObservableProperty]
    private double _llmDownloadPercent;

    [ObservableProperty]
    private bool _isLlmDownloadInProgress;

    [ObservableProperty]
    private string _llmModelPath = string.Empty;

    [ObservableProperty]
    private string _transcriptionLanguage = "en";

    [ObservableProperty]
    private int _indexedPromptSlideCount;

    [ObservableProperty]
    private double _confidenceThreshold = 0.75;

    [ObservableProperty]
    private int _decisionCooldownSeconds = 4;

    [ObservableProperty]
    private int _llmContextSize = 2048;

    [ObservableProperty]
    private int _llmGpuLayerCount;

    [ObservableProperty]
    private bool _isPowerPointConnected;

    [ObservableProperty]
    private bool _isSlideShowRunning;

    [ObservableProperty]
    private bool _isAutomationRunning;

    [ObservableProperty]
    private bool _isAutoAdvanceEnabled = true;

    [ObservableProperty]
    private int _currentSlideNumber;

    [ObservableProperty]
    private PowerPointSlideInfo? _selectedSlide;

    partial void OnIsPowerPointConnectedChanged(bool value) => UpdateCommandStates();

    partial void OnIsSlideShowRunningChanged(bool value) => UpdateCommandStates();

    partial void OnIsAutomationRunningChanged(bool value) => UpdateCommandStates();

    partial void OnIsWhisperDownloadInProgressChanged(bool value) => UpdateCommandStates();

    partial void OnIsLlmDownloadInProgressChanged(bool value) => UpdateCommandStates();

    partial void OnCurrentSlideNumberChanged(int value) => UpdateCommandStates();

    partial void OnSelectedSlideChanged(PowerPointSlideInfo? value) => UpdateCommandStates();

    partial void OnSelectedWhisperDownloadOptionChanged(DownloadableModelOption? value)
    {
        if (value is not null)
        {
            WhisperModelPath = _modelDownloadService.GetStoredFilePath(value.FileName);
            if (!IsWhisperDownloadInProgress)
            {
                UpdateWhisperDownloadStatusFromPath(WhisperModelPath);
            }
        }

        UpdateCommandStates();
    }

    partial void OnWhisperModelPathChanged(string value)
    {
        if (!IsWhisperDownloadInProgress)
        {
            UpdateWhisperDownloadStatusFromPath(value);
        }
    }

    partial void OnLlmDownloadUrlChanged(string value)
    {
        if (Uri.TryCreate(value, UriKind.Absolute, out _))
        {
            LlmModelPath = _modelDownloadService.GetStoredFilePathFromUrl(value);
            if (!IsLlmDownloadInProgress)
            {
                UpdateLlmDownloadStatusFromPath(LlmModelPath);
            }
        }
        else if (!IsLlmDownloadInProgress)
        {
            LlmDownloadPercent = 0;
            LlmDownloadStatus = "Enter a valid direct GGUF download URL.";
        }

        UpdateCommandStates();
    }

    partial void OnLlmModelPathChanged(string value)
    {
        if (!IsLlmDownloadInProgress)
        {
            UpdateLlmDownloadStatusFromPath(value);
        }
    }

    partial void OnIsAutoAdvanceEnabledChanged(bool value)
    {
        StatusMessage = value
            ? "Auto-advance is enabled."
            : "Auto-advance is paused. Transcription will continue until you stop listening.";
    }

    [RelayCommand(CanExecute = nameof(CanDownloadWhisperModel))]
    private async Task DownloadWhisperModelAsync()
    {
        if (SelectedWhisperDownloadOption is null)
        {
            return;
        }

        IsWhisperDownloadInProgress = true;
        WhisperDownloadPercent = 0;
        WhisperDownloadStatus = $"Starting download for {SelectedWhisperDownloadOption.DisplayName}...";
        StatusMessage = WhisperDownloadStatus;

        try
        {
            Progress<ModelDownloadProgress> progress = new(progressUpdate =>
            {
                WhisperDownloadPercent = progressUpdate.PercentComplete;
                WhisperDownloadStatus = progressUpdate.StatusMessage;
            });

            string downloadedPath = await _modelDownloadService
                .DownloadAsync(SelectedWhisperDownloadOption, progress, CancellationToken.None)
                .ConfigureAwait(true);

            WhisperModelPath = downloadedPath;
            WhisperDownloadPercent = 100;
            WhisperDownloadStatus = $"Ready: {Path.GetFileName(downloadedPath)}";
            StatusMessage = $"Downloaded {SelectedWhisperDownloadOption.DisplayName} to {ModelStorageDirectory}.";
        }
        catch (Exception ex)
        {
            WhisperDownloadPercent = 0;
            WhisperDownloadStatus = $"Download failed: {ex.Message}";
            StatusMessage = WhisperDownloadStatus;
        }
        finally
        {
            IsWhisperDownloadInProgress = false;
        }
    }

    [RelayCommand(CanExecute = nameof(CanDownloadLlmModel))]
    private async Task DownloadLlmModelAsync()
    {
        if (!Uri.TryCreate(LlmDownloadUrl, UriKind.Absolute, out _))
        {
            LlmDownloadPercent = 0;
            LlmDownloadStatus = "Enter a valid direct GGUF download URL.";
            StatusMessage = LlmDownloadStatus;
            return;
        }

        IsLlmDownloadInProgress = true;
        LlmDownloadPercent = 0;
        LlmDownloadStatus = "Starting GGUF download...";
        StatusMessage = LlmDownloadStatus;

        try
        {
            Progress<ModelDownloadProgress> progress = new(progressUpdate =>
            {
                LlmDownloadPercent = progressUpdate.PercentComplete;
                LlmDownloadStatus = progressUpdate.StatusMessage;
            });

            string downloadedPath = await _modelDownloadService
                .DownloadFromUrlAsync(LlmDownloadUrl, explicitFileName: null, progress, CancellationToken.None)
                .ConfigureAwait(true);

            LlmModelPath = downloadedPath;
            LlmDownloadPercent = 100;
            LlmDownloadStatus = $"Ready: {Path.GetFileName(downloadedPath)}";
            StatusMessage = $"Downloaded the GGUF model to {ModelStorageDirectory}.";
        }
        catch (Exception ex)
        {
            LlmDownloadPercent = 0;
            LlmDownloadStatus = $"Download failed: {ex.Message}";
            StatusMessage = LlmDownloadStatus;
        }
        finally
        {
            IsLlmDownloadInProgress = false;
        }
    }

    [RelayCommand]
    private void OpenModelFolder()
    {
        try
        {
            Directory.CreateDirectory(ModelStorageDirectory);

            Process.Start(new ProcessStartInfo
            {
                FileName = ModelStorageDirectory,
                UseShellExecute = true,
            });

            StatusMessage = $"Opened {ModelStorageDirectory}.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Could not open the model folder: {ex.Message}";
        }
    }

    [RelayCommand]
    private async Task RefreshPowerPointAsync()
    {
        try
        {
            PresentationSessionSnapshot snapshot = await _powerPointSessionService.ConnectAsync().ConfigureAwait(true);
            ApplySnapshot(snapshot);
            StatusMessage = snapshot.SlideShowState.IsRunning
                ? "PowerPoint connected. The slide show is live."
                : "PowerPoint connected. Start the slide show to enable automation.";
        }
        catch (Exception ex)
        {
            ApplyConnectionFailure(ex.Message);
        }
    }

    [RelayCommand(CanExecute = nameof(CanStartAutomation))]
    private async Task StartAutomationAsync()
    {
        try
        {
            EnsureModelPath(WhisperModelPath, "Whisper model");
            EnsureModelPath(LlmModelPath, "LLM model");

            if (!IsPowerPointConnected)
            {
                await RefreshPowerPointAsync().ConfigureAwait(true);
            }

            if (!IsPowerPointConnected)
            {
                return;
            }

            if (!IsSlideShowRunning)
            {
                throw new InvalidOperationException("Start the PowerPoint slide show before enabling automation.");
            }

            ResetTranscript();
            _lastAdvanceAt = null;

            _automationCts?.Cancel();
            _automationCts?.Dispose();
            _automationCts = new CancellationTokenSource();

            TranscriptionOptions transcriptionOptions = new(
                NormalizeFilePath(WhisperModelPath),
                NormalizeLanguage(TranscriptionLanguage));

            await _transcriptionService.StartAsync(transcriptionOptions, HandleTranscriptChunkAsync, _automationCts.Token).ConfigureAwait(true);

            IsAutomationRunning = true;
            StatusMessage = "Listening to the microphone and evaluating prompts on-device.";
            LastDecision = "Waiting for transcript...";
        }
        catch (Exception ex)
        {
            IsAutomationRunning = false;
            StatusMessage = $"Automation could not start: {ex.Message}";
            LastDecision = ex.Message;
        }
    }

    [RelayCommand(CanExecute = nameof(CanStopAutomation))]
    private async Task StopAutomationAsync()
    {
        _automationCts?.Cancel();
        _automationCts?.Dispose();
        _automationCts = null;

        await _transcriptionService.StopAsync().ConfigureAwait(true);

        IsAutomationRunning = false;
        StatusMessage = "Automation stopped.";
    }

    [RelayCommand(CanExecute = nameof(CanAdvanceSlide))]
    private async Task AdvanceSlideAsync()
    {
        try
        {
            await _powerPointSessionService.AdvanceToNextSlideAsync().ConfigureAwait(true);
            SlideShowState slideShowState = await _powerPointSessionService.GetSlideShowStateAsync().ConfigureAwait(true);
            ApplySlideShowState(slideShowState);
            LastDecision = "Moved forward manually.";
            StatusMessage = $"Moved to slide {slideShowState.CurrentSlideNumber}.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Could not advance the slide: {ex.Message}";
        }
    }

    [RelayCommand(CanExecute = nameof(CanMoveToPreviousSlide))]
    private async Task PreviousSlideAsync()
    {
        if (CurrentSlideNumber <= 1)
        {
            return;
        }

        await MoveToSlideAsync(CurrentSlideNumber - 1, "Moved back manually.").ConfigureAwait(true);
    }

    [RelayCommand(CanExecute = nameof(CanGoToSelectedSlide))]
    private async Task GoToSelectedSlideAsync()
    {
        if (SelectedSlide is null)
        {
            return;
        }

        await MoveToSlideAsync(SelectedSlide.SlideNumber, $"Moved manually to slide {SelectedSlide.SlideNumber}.").ConfigureAwait(true);
    }

    [RelayCommand]
    private void ClearTranscript()
    {
        ResetTranscript();
        LastDecision = "Transcript cleared.";
    }

    private bool CanDownloadWhisperModel() => !IsWhisperDownloadInProgress && SelectedWhisperDownloadOption is not null;

    private bool CanDownloadLlmModel() => !IsLlmDownloadInProgress && Uri.TryCreate(LlmDownloadUrl, UriKind.Absolute, out _);

    private bool CanStartAutomation() => !IsAutomationRunning;

    private bool CanStopAutomation() => IsAutomationRunning;

    private bool CanAdvanceSlide() =>
        IsPowerPointConnected &&
        IsSlideShowRunning &&
        (CurrentSlideNumber <= 0 || _slideCache.Count == 0 || CurrentSlideNumber < _slideCache[^1].SlideNumber);

    private bool CanMoveToPreviousSlide() =>
        IsPowerPointConnected &&
        IsSlideShowRunning &&
        CurrentSlideNumber > 1;

    private bool CanGoToSelectedSlide() =>
        IsPowerPointConnected &&
        IsSlideShowRunning &&
        SelectedSlide is { SlideNumber: > 0 } slide &&
        slide.SlideNumber != CurrentSlideNumber;

    private async Task HandleTranscriptChunkAsync(string transcriptChunk)
    {
        if (string.IsNullOrWhiteSpace(transcriptChunk))
        {
            return;
        }

        await _dispatcher.InvokeAsync(() => AppendTranscriptChunk(transcriptChunk)).Task.ConfigureAwait(false);

        bool shouldEvaluate = await _dispatcher.InvokeAsync(
            () => IsAutomationRunning && IsAutoAdvanceEnabled && IsSlideShowRunning).Task.ConfigureAwait(false);

        if (!shouldEvaluate || !await _evaluationGate.WaitAsync(0).ConfigureAwait(false))
        {
            return;
        }

        try
        {
            await EvaluateCurrentTranscriptWindowAsync(_automationCts?.Token ?? CancellationToken.None).ConfigureAwait(false);
        }
        finally
        {
            _evaluationGate.Release();
        }
    }

    private async Task EvaluateCurrentTranscriptWindowAsync(CancellationToken cancellationToken)
    {
        SlideShowState slideShowState = await _powerPointSessionService.GetSlideShowStateAsync(cancellationToken).ConfigureAwait(false);
        await _dispatcher.InvokeAsync(() => ApplySlideShowState(slideShowState)).Task.ConfigureAwait(false);

        if (!slideShowState.IsRunning)
        {
            await _dispatcher.InvokeAsync(() =>
            {
                StatusMessage = "The slide show is no longer running.";
                LastDecision = "Waiting for the slide show to resume.";
            }).Task.ConfigureAwait(false);
            return;
        }

        string transcript = BuildTranscriptWindow();
        if (string.IsNullOrWhiteSpace(transcript))
        {
            return;
        }

        IReadOnlyList<PowerPointSlideInfo> slideCache = _slideCache;
        if (slideCache.Count == 0)
        {
            return;
        }

        PowerPointSlideInfo currentSlide = slideCache.FirstOrDefault(slide => slide.SlideNumber == slideShowState.CurrentSlideNumber)
            ?? slideCache[0];

        IReadOnlyList<PowerPointSlideInfo> candidateSlides = SlideAutomationIndex.BuildCandidates(_automationPromptIndex, currentSlide.SlideNumber);

        if (candidateSlides.Count == 0)
        {
            await _dispatcher.InvokeAsync(() =>
            {
                LastDecision = "No other indexed slides contain automation prompts.";
            }).Task.ConfigureAwait(false);
            return;
        }

        SlideEvaluationContext evaluationContext = new(currentSlide, candidateSlides, transcript);
        LlmOptions llmOptions = new(
            NormalizeFilePath(LlmModelPath),
            Math.Max(1024, LlmContextSize),
            Math.Max(0, LlmGpuLayerCount));

        SlideSwitchEvaluation evaluation = await _slideEvaluationService
            .EvaluateAsync(evaluationContext, llmOptions, cancellationToken)
            .ConfigureAwait(false);

        AutoAdvanceDecisionOutcome outcome = _autoAdvancePolicy.Evaluate(
            evaluation,
            candidateSlides.Select(slide => slide.SlideNumber).ToHashSet(),
            Math.Clamp(ConfidenceThreshold, 0.0, 1.0),
            DateTimeOffset.UtcNow,
            _lastAdvanceAt,
            TimeSpan.FromSeconds(Math.Max(0, DecisionCooldownSeconds)));

        await _dispatcher.InvokeAsync(() => LastDecision = outcome.Summary).Task.ConfigureAwait(false);

        if (!outcome.ShouldAdvance || outcome.TargetSlideNumber is null)
        {
            return;
        }

        await _powerPointSessionService.GoToSlideAsync(outcome.TargetSlideNumber.Value, cancellationToken).ConfigureAwait(false);

        _lastAdvanceAt = DateTimeOffset.UtcNow;

        SlideShowState updatedState = await _powerPointSessionService.GetSlideShowStateAsync(cancellationToken).ConfigureAwait(false);
        await _dispatcher.InvokeAsync(() =>
        {
            ApplySlideShowState(updatedState);
            StatusMessage = $"Moved automatically to slide {updatedState.CurrentSlideNumber}.";
        }).Task.ConfigureAwait(false);
    }

    private void AppendTranscriptChunk(string transcriptChunk)
    {
        lock (_transcriptLock)
        {
            _transcriptChunks.Enqueue(transcriptChunk.Trim());

            while (_transcriptChunks.Count > MaxTranscriptChunks)
            {
                _transcriptChunks.Dequeue();
            }

            string transcript = string.Join(" ", _transcriptChunks);
            TranscriptWindow = transcript.Length <= MaxTranscriptCharacters
                ? transcript
                : transcript[^MaxTranscriptCharacters..];
        }
    }

    private string BuildTranscriptWindow()
    {
        lock (_transcriptLock)
        {
            return string.Join(" ", _transcriptChunks);
        }
    }

    private void ResetTranscript()
    {
        lock (_transcriptLock)
        {
            _transcriptChunks.Clear();
        }

        TranscriptWindow = string.Empty;
    }

    private void ApplySnapshot(PresentationSessionSnapshot snapshot)
    {
        _slideCache = snapshot.Slides;
        _automationPromptIndex = SlideAutomationIndex.Build(snapshot.Slides);

        Slides.Clear();
        foreach (PowerPointSlideInfo slide in snapshot.Slides)
        {
            Slides.Add(slide);
        }

        PresentationName = snapshot.PresentationName;
        IsPowerPointConnected = true;
        ConnectionSummary = $"{snapshot.PresentationName} loaded with {snapshot.Slides.Count} slides.";
        IndexedPromptSlideCount = _automationPromptIndex.Count;

        SelectedSlide = ResolveSelectedSlide(snapshot.Slides, snapshot.SlideShowState.CurrentSlideNumber);

        ApplySlideShowState(snapshot.SlideShowState);
    }

    private void ApplySlideShowState(SlideShowState slideShowState)
    {
        int previousSlideNumber = CurrentSlideNumber;
        IsSlideShowRunning = slideShowState.IsRunning;
        CurrentSlideNumber = slideShowState.CurrentSlideNumber;

        bool slideChanged = previousSlideNumber > 0 &&
                            slideShowState.CurrentSlideNumber > 0 &&
                            slideShowState.CurrentSlideNumber != previousSlideNumber;

        if (slideChanged)
        {
            ResetTranscript();
            _lastAdvanceAt = DateTimeOffset.UtcNow;

            if (IsAutomationRunning)
            {
                LastDecision = $"Slide changed to {slideShowState.CurrentSlideNumber}. Transcript window reset.";
            }
        }

        SlideShowStateText = slideShowState.IsRunning
            ? $"Running on slide {slideShowState.CurrentSlideNumber}."
            : "Connected, but the slide show is not running.";

        UpdateSlideContext(slideShowState.CurrentSlideNumber);
    }

    private void UpdateSlideContext(int activeSlideNumber)
    {
        PowerPointSlideInfo? currentSlide = null;
        if (_slideCache.Count > 0)
        {
            currentSlide = activeSlideNumber > 0
                ? _slideCache.FirstOrDefault(slide => slide.SlideNumber == activeSlideNumber)
                : _slideCache[0];
        }

        PowerPointSlideInfo? nextSlide = SlideAutomationIndex.FindNearest(_automationPromptIndex, activeSlideNumber);

        CurrentSlideTitle = currentSlide is null
            ? "No active slide"
            : $"Slide {currentSlide.SlideNumber}: {currentSlide.Title}";

        CurrentSlidePrompt = currentSlide is null
            ? "No current slide prompt loaded."
            : string.IsNullOrWhiteSpace(currentSlide.AutomationPrompt)
                ? "This slide does not contain an automation prompt below the --- delimiter."
                : currentSlide.AutomationPrompt;

        NextSlideTitle = nextSlide is null
            ? "No other indexed slide with an automation prompt"
            : $"Nearest indexed prompt: slide {nextSlide.SlideNumber}: {nextSlide.Title}";

        NextSlidePrompt = nextSlide?.AutomationPrompt
            ?? "Add a `---` delimiter in the slide notes and place the machine prompt below it.";
    }

    private void ApplyConnectionFailure(string reason)
    {
        _slideCache = Array.Empty<PowerPointSlideInfo>();
        _automationPromptIndex = Array.Empty<PowerPointSlideInfo>();
        Slides.Clear();
        PresentationName = "Not connected";
        ConnectionSummary = reason;
        SlideShowStateText = "No slide show detected.";
        IndexedPromptSlideCount = 0;
        CurrentSlideNumber = 0;
        SelectedSlide = null;
        CurrentSlideTitle = "No active slide";
        CurrentSlidePrompt = "Open PowerPoint and refresh the session.";
        NextSlideTitle = "No indexed prompt slide";
        NextSlidePrompt = "Refresh after a presentation is open.";
        IsPowerPointConnected = false;
        IsSlideShowRunning = false;
        StatusMessage = $"PowerPoint connection failed: {reason}";
    }

    private void RefreshModelAvailability()
    {
        UpdateWhisperDownloadStatusFromPath(WhisperModelPath);
        UpdateLlmDownloadStatusFromPath(LlmModelPath);
    }

    private void UpdateWhisperDownloadStatusFromPath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            WhisperDownloadPercent = 0;
            WhisperDownloadStatus = "Select or download a Whisper model.";
            return;
        }

        if (File.Exists(path))
        {
            WhisperDownloadPercent = 100;
            WhisperDownloadStatus = $"Ready: {Path.GetFileName(path)}";
            return;
        }

        WhisperDownloadPercent = 0;
        WhisperDownloadStatus = SelectedWhisperDownloadOption is null
            ? "Select or download a Whisper model."
            : $"{SelectedWhisperDownloadOption.DisplayName} will be stored in {ModelStorageDirectory}.";
    }

    private void UpdateLlmDownloadStatusFromPath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            LlmDownloadPercent = 0;
            LlmDownloadStatus = "Paste a direct GGUF URL or use the recommended default.";
            return;
        }

        if (File.Exists(path))
        {
            LlmDownloadPercent = 100;
            LlmDownloadStatus = $"Ready: {Path.GetFileName(path)}";
            return;
        }

        if (!Uri.TryCreate(LlmDownloadUrl, UriKind.Absolute, out _))
        {
            LlmDownloadPercent = 0;
            LlmDownloadStatus = "Enter a valid direct GGUF download URL.";
            return;
        }

        LlmDownloadPercent = 0;
        LlmDownloadStatus = $"The downloaded GGUF will be stored in {ModelStorageDirectory}.";
    }

    private void UpdateCommandStates()
    {
        DownloadWhisperModelCommand.NotifyCanExecuteChanged();
        DownloadLlmModelCommand.NotifyCanExecuteChanged();
        StartAutomationCommand.NotifyCanExecuteChanged();
        StopAutomationCommand.NotifyCanExecuteChanged();
        PreviousSlideCommand.NotifyCanExecuteChanged();
        AdvanceSlideCommand.NotifyCanExecuteChanged();
        GoToSelectedSlideCommand.NotifyCanExecuteChanged();
    }

    private async Task MoveToSlideAsync(int slideNumber, string decisionMessage)
    {
        try
        {
            await _powerPointSessionService.GoToSlideAsync(slideNumber).ConfigureAwait(true);
            SlideShowState slideShowState = await _powerPointSessionService.GetSlideShowStateAsync().ConfigureAwait(true);
            ApplySlideShowState(slideShowState);
            LastDecision = decisionMessage;
            StatusMessage = $"Moved to slide {slideShowState.CurrentSlideNumber}.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Could not move to slide {slideNumber}: {ex.Message}";
        }
    }

    private PowerPointSlideInfo? ResolveSelectedSlide(IReadOnlyList<PowerPointSlideInfo> slides, int currentSlideNumber)
    {
        if (slides.Count == 0)
        {
            return null;
        }

        if (SelectedSlide is not null)
        {
            PowerPointSlideInfo? matchingSelectedSlide = slides.FirstOrDefault(slide => slide.SlideNumber == SelectedSlide.SlideNumber);
            if (matchingSelectedSlide is not null)
            {
                return matchingSelectedSlide;
            }
        }

        return slides.FirstOrDefault(slide => slide.SlideNumber == currentSlideNumber) ?? slides[0];
    }

    private static void EnsureModelPath(string path, string label)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            throw new FileNotFoundException($"{label} file not found.", path);
        }
    }

    private static string NormalizeFilePath(string path) => Path.GetFullPath(path.Trim());

    private static string NormalizeLanguage(string language) =>
        string.IsNullOrWhiteSpace(language) ? "en" : language.Trim();
}
