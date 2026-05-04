using Keboo.VocalSlide.Models;

namespace Keboo.VocalSlide.Services;

public interface IModelDownloadService
{
    string StorageDirectory { get; }

    string GetStoredFilePath(string fileName);

    string GetStoredFilePathFromUrl(string url);

    Task<string> DownloadAsync(
        DownloadableModelOption model,
        IProgress<ModelDownloadProgress>? progress,
        CancellationToken cancellationToken);

    Task<string> DownloadFromUrlAsync(
        string url,
        string? explicitFileName,
        IProgress<ModelDownloadProgress>? progress,
        CancellationToken cancellationToken);
}
