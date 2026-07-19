using Keboo.VocalSlide.Services;
using System.IO;

namespace Keboo.VocalSlide.Tests;

public class ModelCatalogTests
{
    [Test]
    public async Task StorageDirectory_UsesLocalAppData()
    {
        string expected = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Keboo.VocalSlide");

        await Assert.That(ModelCatalog.StorageDirectory).IsEqualTo(expected);
    }

    [Test]
    public async Task GetDownloadFileName_StripsQueryStringFromDirectDownloadUrl()
    {
        string fileName = ModelCatalog.GetDownloadFileName(
            "https://huggingface.co/Qwen/Qwen2.5-0.5B-Instruct-GGUF/resolve/main/qwen2.5-0.5b-instruct-q4_k_m.gguf?download=true");

        await Assert.That(fileName).IsEqualTo("qwen2.5-0.5b-instruct-q4_k_m.gguf");
    }

    [Test]
    public async Task GetStoragePath_PlacesModelsInsideAppDataFolder()
    {
        string path = ModelCatalog.GetStoragePath("ggml-tiny.en.bin");

        await Assert.That(path).IsEqualTo(Path.Combine(ModelCatalog.StorageDirectory, "ggml-tiny.en.bin"));
    }

    [Test]
    public async Task WhisperModels_ExposeCpuAccuracyTiers()
    {
        IReadOnlyList<string> modelIds = ModelCatalog.WhisperModels.Select(model => model.Id).ToArray();

        await Assert.That(modelIds).IsEquivalentTo(
            ["whisper-tiny-en", "whisper-base-en", "whisper-small-en", "whisper-medium-en"]);
        await Assert.That(ModelCatalog.DefaultWhisper.IsCpuRealtimeFriendly).IsTrue();
        await Assert.That(ModelCatalog.MaxAccuracyWhisper.IsCpuRealtimeFriendly).IsFalse();
    }

    [Test]
    public async Task DefaultTranscriptionProfile_MatchesPreviousLatencyDefaults()
    {
        await Assert.That(ModelCatalog.DefaultTranscriptionProfile.Id).IsEqualTo("balanced");
        await Assert.That(ModelCatalog.DefaultTranscriptionProfile.BufferMilliseconds).IsEqualTo(250);
        await Assert.That(ModelCatalog.DefaultTranscriptionProfile.ChunkMilliseconds).IsEqualTo(2000);
    }
}
