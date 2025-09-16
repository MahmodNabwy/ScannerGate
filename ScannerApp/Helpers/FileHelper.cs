using System.Drawing.Imaging;

namespace ScannerApp.Helpers
{
    public static class FileHelper
    {
        public static string GenerateFileName(string path, ImageFormat format)
        {
            var extension = GetExtensionFromFormat(format);
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var uniqueId = Guid.NewGuid().ToString("N")[..8];
            var fileName = $"Scan_{timestamp}_{uniqueId}{extension}";

            return Path.Combine(path, fileName);
        }

        public static string GetExtensionFromFormat(ImageFormat format)
        {
            if (format.Equals(ImageFormat.Jpeg))
                return ".jpg";
            else if (format.Equals(ImageFormat.Png))
                return ".png";
            else if (format.Equals(ImageFormat.Tiff))
                return ".tiff";
            else if (format.Equals(ImageFormat.Bmp))
                return ".bmp";
            else if (format.Equals(ImageFormat.Gif))
                return ".gif";
            else
                return ".jpg"; // Default to JPEG
        }

        public static ImageFormat GetFormatFromExtension(string extension)
        {
            return extension.ToLower() switch
            {
                ".jpg" or ".jpeg" => ImageFormat.Jpeg,
                ".png" => ImageFormat.Png,
                ".tiff" or ".tif" => ImageFormat.Tiff,
                ".bmp" => ImageFormat.Bmp,
                ".gif" => ImageFormat.Gif,
                _ => ImageFormat.Jpeg
            };
        }

        public static void EnsureDirectoryExists(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        public static bool IsValidPath(string path)
        {
            try
            {
                Path.GetFullPath(path);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static long GetFileSize(string filePath)
        {
            if (File.Exists(filePath))
            {
                return new FileInfo(filePath).Length;
            }
            return 0;
        }

        public static string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }
    }
}
