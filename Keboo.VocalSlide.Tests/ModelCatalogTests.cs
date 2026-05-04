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
}
