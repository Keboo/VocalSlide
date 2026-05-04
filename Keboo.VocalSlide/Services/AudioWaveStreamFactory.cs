using System.IO;
using System.Text;

namespace Keboo.VocalSlide.Services;

internal static class AudioWaveStreamFactory
{
    public static MemoryStream CreateWaveStream(byte[] pcmBuffer, int sampleRate, int channels, int bitsPerSample)
    {
        MemoryStream stream = new();
        using BinaryWriter writer = new(stream, Encoding.UTF8, leaveOpen: true);

        int byteRate = sampleRate * channels * bitsPerSample / 8;
        short blockAlign = (short)(channels * bitsPerSample / 8);

        writer.Write(Encoding.ASCII.GetBytes("RIFF"));
        writer.Write(36 + pcmBuffer.Length);
        writer.Write(Encoding.ASCII.GetBytes("WAVE"));
        writer.Write(Encoding.ASCII.GetBytes("fmt "));
        writer.Write(16);
        writer.Write((short)1);
        writer.Write((short)channels);
        writer.Write(sampleRate);
        writer.Write(byteRate);
        writer.Write(blockAlign);
        writer.Write((short)bitsPerSample);
        writer.Write(Encoding.ASCII.GetBytes("data"));
        writer.Write(pcmBuffer.Length);
        writer.Write(pcmBuffer);
        writer.Flush();

        stream.Position = 0;
        return stream;
    }
}
