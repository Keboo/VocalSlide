using Keboo.VocalSlide.Models;
using NAudio.Wave;
using System.IO;
using System.Text;
using Whisper.net;

namespace Keboo.VocalSlide.Services;

public sealed class WhisperTranscriptionService : ILocalTranscriptionService, IAsyncDisposable, IDisposable
{
    private readonly Lock _bufferLock = new();

    private MemoryStream _pcmBuffer = new();
    private WaveInEvent? _waveIn;
    private WhisperFactory? _whisperFactory;
    private Task? _processingTask;
    private CancellationTokenSource? _internalCancellationTokenSource;
    private Func<string, Task>? _onTranscript;

    public bool IsRunning { get; private set; }

    public async Task StartAsync(TranscriptionOptions options, Func<string, Task> onTranscript, CancellationToken cancellationToken)
    {
        if (IsRunning)
        {
            return;
        }

        if (!File.Exists(options.ModelPath))
        {
            throw new FileNotFoundException("Whisper model file not found.", options.ModelPath);
        }

        _onTranscript = onTranscript;
        _internalCancellationTokenSource = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        _whisperFactory = await Task.Run(() => WhisperFactory.FromPath(options.ModelPath), cancellationToken).ConfigureAwait(false);

        _waveIn = new WaveInEvent
        {
            DeviceNumber = 0,
            WaveFormat = new WaveFormat(options.SampleRateHz, options.BitsPerSample, options.Channels),
            BufferMilliseconds = options.BufferMilliseconds,
        };

        _waveIn.DataAvailable += OnDataAvailable;
        _waveIn.StartRecording();

        IsRunning = true;
        _processingTask = Task.Run(() => ProcessAudioAsync(options, _internalCancellationTokenSource.Token), _internalCancellationTokenSource.Token);
    }

    public async Task StopAsync()
    {
        if (!IsRunning)
        {
            return;
        }

        IsRunning = false;

        if (_internalCancellationTokenSource is not null)
        {
            await _internalCancellationTokenSource.CancelAsync().ConfigureAwait(false);
        }

        if (_waveIn is not null)
        {
            _waveIn.DataAvailable -= OnDataAvailable;
            _waveIn.StopRecording();
            _waveIn.Dispose();
            _waveIn = null;
        }

        if (_processingTask is not null)
        {
            try
            {
                await _processingTask.ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                // Expected when the microphone pipeline is stopped.
            }

            _processingTask = null;
        }

        lock (_bufferLock)
        {
            _pcmBuffer.Dispose();
            _pcmBuffer = new MemoryStream();
        }

        _whisperFactory?.Dispose();
        _whisperFactory = null;

        _internalCancellationTokenSource?.Dispose();
        _internalCancellationTokenSource = null;
        _onTranscript = null;
    }

    private void OnDataAvailable(object? sender, WaveInEventArgs e)
    {
        lock (_bufferLock)
        {
            _pcmBuffer.Write(e.Buffer, 0, e.BytesRecorded);
        }
    }

    private async Task ProcessAudioAsync(TranscriptionOptions options, CancellationToken cancellationToken)
    {
        if (_whisperFactory is null)
        {
            return;
        }

        var builder = _whisperFactory.CreateBuilder();
        if (!string.Equals(options.Language, "auto", StringComparison.OrdinalIgnoreCase))
        {
            builder = builder.WithLanguage(options.Language);
        }

        using var processor = builder.Build();

        int minimumBytes = options.SampleRateHz * options.Channels * (options.BitsPerSample / 8) * options.ChunkMilliseconds / 1000;

        while (!cancellationToken.IsCancellationRequested)
        {
            await Task.Delay(options.ChunkMilliseconds, cancellationToken).ConfigureAwait(false);

            byte[]? pcmChunk = null;
            lock (_bufferLock)
            {
                if (_pcmBuffer.Length >= minimumBytes)
                {
                    pcmChunk = _pcmBuffer.ToArray();
                    _pcmBuffer.SetLength(0);
                    _pcmBuffer.Position = 0;
                }
            }

            if (pcmChunk is null || pcmChunk.Length == 0)
            {
                continue;
            }

            using MemoryStream wavStream = AudioWaveStreamFactory.CreateWaveStream(
                pcmChunk,
                options.SampleRateHz,
                options.Channels,
                options.BitsPerSample);

            StringBuilder transcriptBuilder = new();
            await foreach (var segment in processor.ProcessAsync(wavStream, cancellationToken).ConfigureAwait(false))
            {
                if (segment.NoSpeechProbability > 0.8) continue;
                if (!string.IsNullOrWhiteSpace(segment.Text))
                {
                    transcriptBuilder.Append(segment.Text.Trim()).Append(' ');
                }
            }

            string transcript = transcriptBuilder.ToString().Trim();
            if (!string.IsNullOrWhiteSpace(transcript) && _onTranscript is not null)
            {
                await _onTranscript(transcript).ConfigureAwait(false);
            }
        }
    }

    public async ValueTask DisposeAsync()
    {
        await StopAsync().ConfigureAwait(false);
    }

    public void Dispose()
    {
        StopAsync().GetAwaiter().GetResult();
    }
}
