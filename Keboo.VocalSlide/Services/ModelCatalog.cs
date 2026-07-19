using Keboo.VocalSlide.Models;
using System.IO;

namespace Keboo.VocalSlide.Services;

public static class ModelCatalog
{
    public static DownloadableModelOption DefaultWhisper { get; } = new(
        "whisper-tiny-en",
        "Whisper Tiny English (Recommended for speed)",
        "ggml-tiny.en.bin",
        "https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-tiny.en.bin?download=true",
        "Fastest English whisper.cpp model and the default choice for low latency.");

    public static DownloadableModelOption BalancedWhisper { get; } = new(
        "whisper-base-en",
        "Whisper Base English (Balanced)",
        "ggml-base.en.bin",
        "https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-base.en.bin?download=true",
        "Higher-quality English transcription with a larger runtime cost.");

    public static DownloadableModelOption AccurateWhisper { get; } = new(
        "whisper-small-en",
        "Whisper Small English (Higher accuracy)",
        "ggml-small.en.bin",
        "https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-small.en.bin?download=true",
        "Improved English transcription quality with a noticeable CPU cost.");

    public static DownloadableModelOption MaxAccuracyWhisper { get; } = new(
        "whisper-medium-en",
        "Whisper Medium English (High accuracy, slower)",
        "ggml-medium.en.bin",
        "https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-medium.en.bin?download=true",
        "Highest accuracy option in this app, but often too slow for strict real-time CPU usage.",
        IsCpuRealtimeFriendly: false);

    public static IReadOnlyList<DownloadableModelOption> WhisperModels { get; } =
        [DefaultWhisper, BalancedWhisper, AccurateWhisper, MaxAccuracyWhisper];

    public static TranscriptionProfileOption DefaultTranscriptionProfile { get; } = new(
        "balanced",
        "Balanced (Recommended)",
        BufferMilliseconds: 250,
        ChunkMilliseconds: 2000,
        "Default profile. Good tradeoff between latency and transcript quality for CPU use.");

    public static TranscriptionProfileOption LowLatencyTranscriptionProfile { get; } = new(
        "low-latency",
        "Low latency",
        BufferMilliseconds: 150,
        ChunkMilliseconds: 1200,
        "Faster updates and lower delay, with slightly lower recognition quality.");

    public static TranscriptionProfileOption AccuracyTranscriptionProfile { get; } = new(
        "high-accuracy",
        "Higher accuracy",
        BufferMilliseconds: 300,
        ChunkMilliseconds: 3000,
        "More context per pass for better recognition, with additional lag.");

    public static IReadOnlyList<TranscriptionProfileOption> TranscriptionProfiles { get; } =
        [LowLatencyTranscriptionProfile, DefaultTranscriptionProfile, AccuracyTranscriptionProfile];

    public static string DefaultOllamaModelName { get; } = "qwen2.5:0.5b";

    public static string StorageDirectory { get; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Keboo.VocalSlide");

    public static string GetStoragePath(string fileName) =>
        Path.Combine(StorageDirectory, SanitizeFileName(fileName));

    public static string GetDownloadFileName(string url)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(url);

        if (!Uri.TryCreate(url, UriKind.Absolute, out Uri? uri))
        {
            throw new ArgumentException("The model download URL must be absolute.", nameof(url));
        }

        string fileName = Path.GetFileName(Uri.UnescapeDataString(uri.AbsolutePath));
        if (string.IsNullOrWhiteSpace(fileName))
        {
            throw new ArgumentException("Could not infer a file name from the model download URL.", nameof(url));
        }

        return SanitizeFileName(fileName);
    }

    private static string SanitizeFileName(string fileName)
    {
        string sanitized = fileName.Trim();

        foreach (char invalidCharacter in Path.GetInvalidFileNameChars())
        {
            sanitized = sanitized.Replace(invalidCharacter, '_');
        }

        return sanitized;
    }
}
