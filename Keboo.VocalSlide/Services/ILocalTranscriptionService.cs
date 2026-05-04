using Keboo.VocalSlide.Models;

namespace Keboo.VocalSlide.Services;

public interface ILocalTranscriptionService
{
    bool IsRunning { get; }

    Task StartAsync(TranscriptionOptions options, Func<string, Task> onTranscript, CancellationToken cancellationToken);

    Task StopAsync();
}
