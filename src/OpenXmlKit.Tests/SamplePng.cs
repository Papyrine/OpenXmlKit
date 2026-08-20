using System.IO.Compression;

/// <summary>
/// Builds a real PNG of a given size, so the image tests assert against something Word would
/// actually open rather than against a plausible header.
/// </summary>
public static class SamplePng
{
    public static byte[] Create(int width, int height)
    {
        using var stream = new MemoryStream();
        stream.Write([0x89, (byte) 'P', (byte) 'N', (byte) 'G', 0x0D, 0x0A, 0x1A, 0x0A], 0, 8);

        var header = new byte[13];
        WriteBigEndian(header, 0, width);
        WriteBigEndian(header, 4, height);
        header[8] = 8; // bit depth
        header[9] = 2; // truecolour
        WriteChunk(stream, "IHDR", header);

        // One filter byte per scanline, then three bytes a pixel. All zero, so the image is black.
        var raw = new byte[height * (1 + width * 3)];
        using (var compressed = new MemoryStream())
        {
            using (var deflate = new ZLibStream(compressed, CompressionLevel.Optimal, true))
            {
                deflate.Write(raw, 0, raw.Length);
            }

            WriteChunk(stream, "IDAT", compressed.ToArray());
        }

        WriteChunk(stream, "IEND", []);
        return stream.ToArray();
    }

    static void WriteChunk(Stream stream, string type, byte[] data)
    {
        var length = new byte[4];
        WriteBigEndian(length, 0, data.Length);
        stream.Write(length, 0, 4);

        var typeBytes = Encoding.ASCII.GetBytes(type);
        stream.Write(typeBytes, 0, typeBytes.Length);
        stream.Write(data, 0, data.Length);

        var crc = new byte[4];
        WriteBigEndian(crc, 0, (int) Crc32([.. typeBytes, .. data]));
        stream.Write(crc, 0, 4);
    }

    static void WriteBigEndian(byte[] target, int offset, int value)
    {
        target[offset] = (byte) (value >> 24);
        target[offset + 1] = (byte) (value >> 16);
        target[offset + 2] = (byte) (value >> 8);
        target[offset + 3] = (byte) value;
    }

    static uint Crc32(byte[] bytes)
    {
        var crc = 0xFFFFFFFFu;
        foreach (var b in bytes)
        {
            crc ^= b;
            for (var bit = 0; bit < 8; bit++)
            {
                crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320u : crc >> 1;
            }
        }

        return crc ^ 0xFFFFFFFFu;
    }
}
