using Keboo.VocalSlide.Models;
using System.IO;
using System.Net.Http;

namespace Keboo.VocalSlide.Services;

public sealed class ModelDownloadService : IModelDownloadService, IDisposable
{
    private readonly HttpClient _httpClient;

    public ModelDownloadService()
    {
        _httpClient = new HttpClient();
        _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("KebooVocalSlide/1.0");

        Directory.CreateDirectory(StorageDirectory);
    }

    public string StorageDirectory => ModelCatalog.StorageDirectory;

    public string GetStoredFilePath(string fileName) => ModelCatalog.GetStoragePath(fileName);

    public string GetStoredFilePathFromUrl(string url) => ModelCatalog.GetStoragePath(ModelCatalog.GetDownloadFileName(url));

    public Task<string> DownloadAsync(
        DownloadableModelOption model,
        IProgress<ModelDownloadProgress>? progress,
        CancellationToken cancellationToken) =>
        DownloadFromUrlAsync(model.DownloadUrl, model.FileName, progress, cancellationToken);

    public async Task<string> DownloadFromUrlAsync(
        string url,
        string? explicitFileName,
        IProgress<ModelDownloadProgress>? progress,
        CancellationToken cancellationToken)
    {
        string fileName = string.IsNullOrWhiteSpace(explicitFileName)
            ? ModelCatalog.GetDownloadFileName(url)
            : explicitFileName.Trim();

        string targetPath = GetStoredFilePath(fileName);
        string tempPath = $"{targetPath}.download";

        Directory.CreateDirectory(StorageDirectory);

        FileInfo existingFile = new(targetPath);
        if (existingFile.Exists && existingFile.Length > 0)
        {
            progress?.Report(ModelDownloadProgress.Completed(fileName, targetPath));
            return targetPath;
        }

        if (existingFile.Exists)
        {
            existingFile.Delete();
        }

        if (File.Exists(tempPath))
        {
            File.Delete(tempPath);
        }

        progress?.Report(ModelDownloadProgress.Starting(fileName, targetPath, $"Starting download for {fileName}..."));

        try
        {
            using HttpResponseMessage response = await _httpClient
                .GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken)
                .ConfigureAwait(false);

            response.EnsureSuccessStatusCode();

            long? totalBytes = response.Content.Headers.ContentLength;

            byte[] buffer = new byte[81_920];
            long bytesDownloaded = 0;

            await using (Stream source = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false))
            await using (FileStream destination = new(tempPath, FileMode.Create, FileAccess.Write, FileShare.None, 81_920, useAsync: true))
            {
                while (true)
                {
                    int read = await source.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken).ConfigureAwait(false);
                    if (read == 0)
                    {
                        break;
                    }

                    await destination.WriteAsync(buffer.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
                    bytesDownloaded += read;

                    string statusMessage = totalBytes is > 0
                        ? $"Downloading {fileName} ({bytesDownloaded / 1024d / 1024d:0.0} MB of {totalBytes.Value / 1024d / 1024d:0.0} MB)"
                        : $"Downloading {fileName} ({bytesDownloaded / 1024d / 1024d:0.0} MB)";

                    progress?.Report(new ModelDownloadProgress(fileName, targetPath, bytesDownloaded, totalBytes, statusMessage));
                }

                await destination.FlushAsync(cancellationToken).ConfigureAwait(false);
            }

            File.Move(tempPath, targetPath, overwrite: true);

            progress?.Report(ModelDownloadProgress.Completed(fileName, targetPath));
            return targetPath;
        }
        catch
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }

            throw;
        }
    }

    public void Dispose()
    {
        _httpClient.Dispose();
    }
}
