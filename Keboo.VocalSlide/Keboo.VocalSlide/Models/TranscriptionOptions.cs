namespace Keboo.VocalSlide.Models;

public sealed record TranscriptionOptions(
    string ModelPath,
    string Language,
    int SampleRateHz = 16000,
    int Channels = 1,
    int BitsPerSample = 16,
    int BufferMilliseconds = 250,
    int ChunkMilliseconds = 2000);
