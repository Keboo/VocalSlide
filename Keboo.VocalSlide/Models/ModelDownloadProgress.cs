namespace Keboo.VocalSlide.Models;

public sealed record ModelDownloadProgress(
    string FileName,
    string DestinationPath,
    long BytesDownloaded,
    long? TotalBytes,
    string StatusMessage)
{
    public double PercentComplete =>
        TotalBytes is > 0
            ? Math.Clamp(BytesDownloaded * 100d / TotalBytes.Value, 0d, 100d)
            : 0d;

    public static ModelDownloadProgress Starting(string fileName, string destinationPath, string statusMessage) =>
        new(fileName, destinationPath, 0, null, statusMessage);

    public static ModelDownloadProgress Completed(string fileName, string destinationPath) =>
        new(fileName, destinationPath, 1, 1, $"Ready: {fileName}");
}
