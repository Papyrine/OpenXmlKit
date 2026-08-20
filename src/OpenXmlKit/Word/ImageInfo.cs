namespace OpenXmlKit.Word;

/// <summary>
/// What can be read about an image from its own bytes: the format, and the size it wants to be.
/// </summary>
/// <remarks>
/// Enough header parsing to answer "how big is this", so that inserting an image does not oblige
/// the caller to already know. Word needs an explicit extent on every drawing — there is no
/// intrinsic sizing in the format — so without this every call site has to carry a width and a
/// height, and the usual result is a picture stretched to whatever numbers were handy.
/// </remarks>
public readonly record struct ImageInfo(ImageFormat Format, int WidthPixels, int HeightPixels, double Dpi)
{
    /// <summary>
    /// The size the image asks to be drawn at, its pixel dimensions taken at its own resolution.
    /// </summary>
    public Length Width => Length.FromPixels(WidthPixels, Dpi);

    public Length Height => Length.FromPixels(HeightPixels, Dpi);

    /// <summary>
    /// Reads the format and dimensions from the image's own header.
    /// </summary>
    public static ImageInfo Read(byte[] bytes)
    {
        if (!TryRead(bytes, out var info))
        {
            throw new NotSupportedException(
                "Could not read the image dimensions. Supported formats are PNG, JPEG, GIF and BMP; " +
                "for anything else, pass an explicit width and height.");
        }

        return info;
    }

    public static bool TryRead(byte[] bytes, out ImageInfo info)
    {
        info = default;
        if (bytes.Length < 24)
        {
            return false;
        }

        if (IsPng(bytes))
        {
            // IHDR is the first chunk and starts at byte 8: length, type, then width and height as
            // big-endian 32-bit values.
            info = new(ImageFormat.Png, BigEndian32(bytes, 16), BigEndian32(bytes, 20), defaultDpi);
            return true;
        }

        if (bytes[0] == 'G' &&
            bytes[1] == 'I' &&
            bytes[2] == 'F')
        {
            info = new(ImageFormat.Gif, LittleEndian16(bytes, 6), LittleEndian16(bytes, 8), defaultDpi);
            return true;
        }

        if (bytes[0] == 'B' &&
            bytes[1] == 'M')
        {
            info = new(ImageFormat.Bmp, LittleEndian32(bytes, 18), Math.Abs(LittleEndian32(bytes, 22)), defaultDpi);
            return true;
        }

        return TryReadJpeg(bytes, out info);
    }

    const double defaultDpi = 96;

    static bool IsPng(byte[] bytes) =>
        bytes[0] == 0x89 &&
        bytes[1] == 'P' &&
        bytes[2] == 'N' &&
        bytes[3] == 'G';

    // JPEG carries its dimensions in a start-of-frame marker somewhere after the header, so the
    // segments have to be walked rather than indexed.
    static bool TryReadJpeg(byte[] bytes, out ImageInfo info)
    {
        info = default;
        if (bytes[0] != 0xFF ||
            bytes[1] != 0xD8)
        {
            return false;
        }

        var index = 2;
        while (index + 9 < bytes.Length)
        {
            if (bytes[index] != 0xFF)
            {
                index++;
                continue;
            }

            var marker = bytes[index + 1];
            // Every SOFn except the four that are not frame headers carries height then width.
            if (marker is >= 0xC0 and <= 0xCF and not (0xC4 or 0xC8 or 0xCC or 0xD8))
            {
                var height = BigEndian16(bytes, index + 5);
                var width = BigEndian16(bytes, index + 7);
                info = new(ImageFormat.Jpeg, width, height, defaultDpi);
                return true;
            }

            var length = BigEndian16(bytes, index + 2);
            if (length < 2)
            {
                return false;
            }

            index += 2 + length;
        }

        return false;
    }

    static int BigEndian32(byte[] bytes, int offset) =>
        (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

    static int BigEndian16(byte[] bytes, int offset) =>
        (bytes[offset] << 8) | bytes[offset + 1];

    static int LittleEndian32(byte[] bytes, int offset) =>
        bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24);

    static int LittleEndian16(byte[] bytes, int offset) =>
        bytes[offset] | (bytes[offset + 1] << 8);
}
