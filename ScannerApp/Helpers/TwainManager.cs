using NTwain.Data;
using NTwain;
using ScannerApp.Models;
using System.Windows.Forms;
using System.Drawing.Imaging;
using ColorMode = ScannerApp.Models.ColorMode;

namespace ScannerApp.Helpers
{
    public class TwainManager : IDisposable
    {
        private readonly string _saveDir = @"C:\Scans";

        private TwainSession? _session;
        private readonly object _lockObject = new object();
        private bool _disposed = false;

        public TwainManager()
        {
            try
            {

            }
            catch (Exception ex)
            {
                File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: failed: Failed to initialize TWAIN session: {ex.Message} {Environment.NewLine}");

                System.Diagnostics.Debug.WriteLine($"Failed to initialize TWAIN session: {ex.Message}");
            }
        }

        public List<ScannerInfo> GetScanners()
        {
            var scanners = new List<ScannerInfo>();
            var appId = TWIdentity.CreateFromAssembly(DataGroups.Image, typeof(ScannerForm).Assembly);
            _session = new TwainSession(appId);
            _session.Open();
            try
            {
                if (_session == null || _session.State < 3)
                {
                    return scanners;
                }

                foreach (var source in _session.GetSources())
                {
                    try
                    {
                        if (IsScannerConnected(source))
                        {
                            var scanner = new ScannerInfo
                            {
                                Id = source.Name,
                                Name = source.Name,
                                Type = ScannerType.TWAIN,
                                SupportsDuplex = false // Default to false
                            };

                            // Check duplex support only once
                            try
                            {
                                source.Open();
                                scanner.SupportsDuplex = CheckDuplexSupport(source) == NTwain.Data.BoolType.True;
                                source.Close();
                            }
                            catch (Exception ex)
                            {
                                File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: Failed to check duplex support for {source.Name}: {ex.Message} {Environment.NewLine}");
                                scanner.SupportsDuplex = false;

                                // Ensure source is closed on error
                                try
                                {
                                    // Attempt to close the source if it is open.
                                    if (source.IsOpen)
                                    {
                                        source.Close();
                                    }
                                }
                                catch (Exception closeEx)
                                {
                                    // Exception ignored intentionally: closing a TWAIN source may fail if the source is already closed or in an invalid state.
                                    // See S2486: Exception is intentionally ignored as a best-effort cleanup.
                                }
                            }

                            scanners.Add(scanner);
                        }
                    }
                    catch (Exception ex)
                    {
                        File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: Failed to process scanner {source.Name}: {ex.Message} {Environment.NewLine}");
                        System.Diagnostics.Debug.WriteLine($"Failed to process scanner {source.Name}: {ex.Message}");
                    }
                }
                if (_session is not null)
                {
                    _session.Close();
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: Failed to enumerate TWAIN scanners: {ex.Message} {Environment.NewLine}");
                System.Diagnostics.Debug.WriteLine($"Failed to enumerate TWAIN scanners: {ex.Message}");
            }

            return scanners;
        }



        public async Task<List<ScanResult>> ScanAsync(ScannerInfo scannerInfo, ScanSettings settings)
        {
            var results = new List<ScanResult>();
            var tcs = new TaskCompletionSource<List<ScanResult>>();
            var appId = TWIdentity.CreateFromAssembly(DataGroups.Image, typeof(ScannerForm).Assembly);
            _session = new TwainSession(appId);

            try
            {
                _session.Open();
                if (_session.State < 3)
                {
                    results.Add(new ScanResult { Success = false, ErrorMessage = "TWAIN session failed to open." });
                    return results;
                }

                var source = _session.GetSources().FirstOrDefault(s => s.Name == scannerInfo.Name);
                if (source == null)
                {
                    results.Add(new ScanResult { Success = false, ErrorMessage = "Scanner source not found." });
                    return results;
                }

                source.Open();
                if (_session.State < 4)
                {
                    results.Add(new ScanResult { Success = false, ErrorMessage = "Scanner source failed to open." });
                    return results;
                }

                bool borderDetectionEnabled = ConfigureScanner(source, settings, scannerInfo);

                int imageCount = 0;
                string frontPath = string.Empty, backPath = string.Empty;
                string sessionId = Guid.NewGuid().ToString();

                _session.DataTransferred += (s, e) =>
                {
                    try
                    {
                        if (e.NativeData != null)
                        {
                            using var stream = e.GetNativeImageStream();
                            if (stream != null)
                            {
                                imageCount++;
                                var imgName = (imageCount == 1) ? "back" : "front";
                                string filePath = Path.Combine(_saveDir, $"{sessionId}_{imgName}.jpg");

                                using (var bmp = new System.Drawing.Bitmap(stream))
                                {
                                    Bitmap imageToSave = bmp;
                                    // If automatic border detection is not supported/enabled, crop manually.
                                    //if (borderDetectionEnabled)
                                    //{
                                    //    imageToSave = CropImage(bmp);
                                    //}

                                    var jpegEncoder = GetEncoder(ImageFormat.Jpeg);
                                    var encoderParameters = new EncoderParameters(1);
                                    encoderParameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 100L); // 95% quality
                                    imageToSave.Save(filePath, jpegEncoder, encoderParameters);

                                    // If a new bitmap was created for cropping, dispose of it.
                                    if (!ReferenceEquals(bmp, imageToSave))
                                    {
                                        imageToSave.Dispose();
                                    }
                                }

                                if (imageCount == 1) frontPath = filePath;
                                if (imageCount == 2) backPath = filePath;

                                results.Add(new ScanResult
                                {
                                    Success = true,
                                    FilePath = filePath,
                                    ScannedAt = DateTime.Now
                                });
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: Failed to save image: {ex.Message}{Environment.NewLine}");
                        results.Add(new ScanResult { Success = false, ErrorMessage = $"Failed to save image: {ex.Message}" });
                    }
                };

                _session.SourceDisabled += (s, e) =>
                {
                    // You can add final logic here if needed, for example, combining images.
                    // For now, we just complete the task.
                    tcs.TrySetResult(results);
                };

                _session.TransferError += (s, e) =>
                {
                    results.Add(new ScanResult { Success = false, ErrorMessage = $"Transfer error: {e.Exception?.Message}" });
                    tcs.TrySetResult(results);
                };

                var enableMode = settings.ShowUI ? SourceEnableMode.ShowUI : SourceEnableMode.NoUI;
                if (source.Enable(enableMode, true, IntPtr.Zero) != ReturnCode.Success)
                {
                    results.Add(new ScanResult { Success = false, ErrorMessage = "Failed to enable scanner." });
                    return results;
                }

                return await tcs.Task;
            }
            catch (Exception ex)
            {
                File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: TWAIN scan failed: {ex.Message} {Environment.NewLine}");
                results.Add(new ScanResult { Success = false, ErrorMessage = $"TWAIN scan failed: {ex.Message}" });
                return results;
            }
            finally
            {
                if (_session != null && _session.State > 2)
                {
                    _session.Close();
                }
            }
        }


        private bool IsScannerConnected(DataSource source)
        {
            bool isConnected = false;
            try
            {
                source.Open();
                var onlineCapability = source.Capabilities.CapDeviceOnline;
                var onlineResult = onlineCapability.GetCurrent();

                isConnected = onlineResult == NTwain.Data.BoolType.True;

                if (isConnected)
                {
                    System.Diagnostics.Debug.WriteLine($"Scanner {source.Name} is online");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"Scanner {source.Name} is offline");
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: Scanner connection test failed for {source.Name}: {ex.Message} {Environment.NewLine}");
                System.Diagnostics.Debug.WriteLine($"Scanner connection test failed for {source.Name}: {ex.Message}");
                isConnected = false;
            }
            finally
            {
                // Ensure source is closed after the check.
                if (source.IsOpen)
                {
                    source.Close();
                }
            }
            return isConnected;
        }


        private NTwain.Data.BoolType CheckDuplexSupport(DataSource source)
        {
            try
            {
                // Assume source is already opened by caller
                var cap = source.Capabilities.CapDuplexEnabled;
                if (cap == null || !cap.IsSupported)
                {
                    return NTwain.Data.BoolType.False;
                }

                // Check if duplex can be enabled
                var supportedCaps = cap.GetValues();
                return supportedCaps.Contains(NTwain.Data.BoolType.True) ? NTwain.Data.BoolType.True : NTwain.Data.BoolType.False;
            }
            catch (Exception ex)
            {
                File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: Failed to check duplex support: {ex.Message} {Environment.NewLine}");
                System.Diagnostics.Debug.WriteLine($"Failed to check duplex support: {ex.Message}");
                return NTwain.Data.BoolType.False;
            }
        }

        // Helper method to handle data transfer events
        private void OnDataTransferred(object sender, DataTransferredEventArgs e)
        {
            // This is just a placeholder - the actual handler is defined inline above
        }


        private ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }

        private bool ConfigureScanner(DataSource ds, ScanSettings settings, ScannerInfo scannerInfo)
        {
            bool borderDetectionSet = false;
            try
            {
                File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: Configuring scanner {ds.Name} {Environment.NewLine}");

                // Try to enable automatic border detection
                try
                {
                    var borderCap = ds.Capabilities.ICapAutomaticBorderDetection;
                    if (borderCap != null && borderCap.IsSupported)
                    {
                        var result = borderCap.SetValue(NTwain.Data.BoolType.True);
                        if (result == ReturnCode.Success)
                        {
                            borderDetectionSet = true;
                        }
                        File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  Automatic Border Detection enabled: {result} {Environment.NewLine}");
                    }
                }
                catch (Exception ex)
                {
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  Failed to set automatic border detection: {ex.Message} {Environment.NewLine}");
                }


                // Set compression to none for highest quality transfer
                try
                {
                    var compressionResult = ds.Capabilities.ICapCompression?.SetValue(CompressionType.None);
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  Compression set to None: {compressionResult} {Environment.NewLine}");
                }
                catch (Exception ex)
                {
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  Failed to set compression: {ex.Message} {Environment.NewLine}");
                }

                // Set pixel type (color mode)
                try
                {
                    var pixelType = settings.ColorMode switch
                    {
                        ColorMode.BlackAndWhite => PixelType.BlackWhite,
                        ColorMode.Grayscale => PixelType.Gray,
                        ColorMode.Color => PixelType.RGB,
                        _ => PixelType.RGB
                    };
                    var pixelResult = ds.Capabilities.ICapPixelType?.SetValue(pixelType);
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  PixelType set to {pixelType}: {pixelResult} {Environment.NewLine}");

                    // Set bit depth for better quality
                    ushort bitDepth = pixelType switch
                    {
                        PixelType.RGB => 24,
                        PixelType.Gray => 8,
                        PixelType.BlackWhite => 1,
                        _ => 24
                    };
                    var bitDepthResult = ds.Capabilities.ICapBitDepth?.SetValue(bitDepth);
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  Bit depth set to {bitDepth}: {bitDepthResult} {Environment.NewLine}");
                }
                catch (Exception ex)
                {
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  Failed to set pixel type or bit depth: {ex.Message} {Environment.NewLine}");
                }

                // Set resolution
                try
                {
                    var xResResult = ds.Capabilities.ICapXResolution?.SetValue((float)settings.Resolution);
                    var yResResult = ds.Capabilities.ICapYResolution?.SetValue((float)settings.Resolution);
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  Resolution set to {settings.Resolution}: X={xResResult}, Y={yResResult} {Environment.NewLine}");
                }
                catch (Exception ex)
                {
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  Failed to set resolution: {ex.Message} {Environment.NewLine}");
                }

                // Set duplex if supported and requested
                try
                {
                    if (scannerInfo.SupportsDuplex)
                    {
                        var duplexResult = ds.Capabilities.CapDuplexEnabled?.SetValue(NTwain.Data.BoolType.True);
                        File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  Duplex enabled: {duplexResult} {Environment.NewLine}");
                    }
                    else
                    {
                        var duplexResult = ds.Capabilities.CapDuplexEnabled?.SetValue(NTwain.Data.BoolType.False);
                        File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  Duplex disabled: {duplexResult} {Environment.NewLine}");
                    }
                }
                catch (Exception ex)
                {
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  Failed to set duplex: {ex.Message} {Environment.NewLine}");
                }

                // Set feeder capabilities
                try
                {
                    var feederResult = ds.Capabilities.CapFeederEnabled?.SetValue(NTwain.Data.BoolType.True);
                    var autoFeedResult = ds.Capabilities.CapAutoFeed?.SetValue(NTwain.Data.BoolType.True);
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  Feeder settings: Feeder={feederResult}, AutoFeed={autoFeedResult} {Environment.NewLine}");

                    // Set transfer count to -1 to scan all available pages
                    var xferCountResult = ds.Capabilities.CapXferCount?.SetValue(-1);
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  Transfer count set to -1: {xferCountResult} {Environment.NewLine}");
                }
                catch (Exception ex)
                {
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  Failed to set feeder capabilities: {ex.Message} {Environment.NewLine}");
                }

                // Disable UI if not requested
                try
                {
                    if (!settings.ShowUI)
                    {
                        var indicatorResult = ds.Capabilities.CapIndicators?.SetValue(NTwain.Data.BoolType.False);
                        File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  Indicators disabled: {indicatorResult} {Environment.NewLine}");
                    }
                }
                catch (Exception ex)
                {
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"  Failed to set UI indicators: {ex.Message} {Environment.NewLine}");
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: Failed to configure scanner: {ds.Name} {Environment.NewLine}");
                System.Diagnostics.Debug.WriteLine($"Failed to configure scanner: {ex.Message}");
            }
            return borderDetectionSet;
        }

        private Bitmap CropImage(Bitmap source, int tolerance = 50)
        {
            // Find the bounding box of the content
            int top = -1, bottom = -1, left = -1, right = -1;

            // Find top
            for (int y = 0; y < source.Height; y++)
            {
                for (int x = 0; x < source.Width; x++)
                {
                    Color pixel = source.GetPixel(x, y);
                    if (pixel.R < 255 - tolerance || pixel.G < 255 - tolerance || pixel.B < 255 - tolerance)
                    {
                        top = y;
                        goto FindBottom;
                    }
                }
            }
        FindBottom:
            if (top == -1) return source; // Image is completely white

            // Find bottom
            for (int y = source.Height - 1; y >= 0; y--)
            {
                for (int x = 0; x < source.Width; x++)
                {
                    Color pixel = source.GetPixel(x, y);
                    if (pixel.R < 255 - tolerance || pixel.G < 255 - tolerance || pixel.B < 255 - tolerance)
                    {
                        bottom = y;
                        goto FindLeft;
                    }
                }
            }
        FindLeft:
            // Find left
            for (int x = 0; x < source.Width; x++)
            {
                for (int y = top; y <= bottom; y++)
                {
                    Color pixel = source.GetPixel(x, y);
                    if (pixel.R < 255 - tolerance || pixel.G < 255 - tolerance || pixel.B < 255 - tolerance)
                    {
                        left = x;
                        goto FindRight;
                    }
                }
            }
        FindRight:
            // Find right
            for (int x = source.Width - 1; x >= 0; x--)
            {
                for (int y = top; y <= bottom; y++)
                {
                    Color pixel = source.GetPixel(x, y);
                    if (pixel.R < 255 - tolerance || pixel.G < 255 - tolerance || pixel.B < 255 - tolerance)
                    {
                        right = x;
                        goto Crop;
                    }
                }
            }

        Crop:
            int width = right - left + 1;
            int height = bottom - top + 1;

            if (width <= 0 || height <= 0) return source; // No content found

            var cropRect = new Rectangle(left, top, width, height);
            Bitmap croppedImage = new Bitmap(cropRect.Width, cropRect.Height);

            using (Graphics g = Graphics.FromImage(croppedImage))
            {
                g.DrawImage(source, new Rectangle(0, 0, croppedImage.Width, croppedImage.Height),
                                 cropRect,
                                 GraphicsUnit.Pixel);
            }

            return croppedImage;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    try
                    {
                        _session?.Close();
                        //_session?.Dispose();
                    }
                    catch (Exception ex)
                    {
                        File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: failed: Error disposing TWAIN session: {ex.Message} {Environment.NewLine}");

                        System.Diagnostics.Debug.WriteLine($"Error disposing TWAIN session: {ex.Message}");
                    }
                }
                _disposed = true;
            }
        }
    }
}
