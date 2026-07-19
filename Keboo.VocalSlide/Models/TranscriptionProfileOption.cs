namespace Keboo.VocalSlide.Models;

public sealed record TranscriptionProfileOption(
    string Id,
    string DisplayName,
    int BufferMilliseconds,
    int ChunkMilliseconds,
    string Description);
