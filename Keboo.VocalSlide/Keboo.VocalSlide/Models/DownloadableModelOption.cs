namespace Keboo.VocalSlide.Models;

public sealed record DownloadableModelOption(
    string Id,
    string DisplayName,
    string FileName,
    string DownloadUrl,
    string Description);
