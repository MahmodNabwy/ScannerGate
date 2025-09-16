using NTwain.Data;
using NTwain;
using ScannerApp.Models;
namespace ScannerApp.Helpers
{
    public class TwainManager : IDisposable
    {
        private TwainSession? _session;
        private readonly object _lockObject = new object();
        private bool _disposed = false;

        public TwainManager()
        {
            try
            {
                _session = new TwainSession(TWIdentity.CreateFromAssembly(DataGroups.Image, System.Reflection.Assembly.GetExecutingAssembly()));
                _session.Open();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to initialize TWAIN session: {ex.Message}");
            }
        }

        public List<ScannerInfo> GetScanners()
        {
            var scanners = new List<ScannerInfo>();

            try
            {
                if (_session == null || _session.State < 3)
                {
                    return scanners;
                }

                foreach (var source in _session.GetSources())
                {
                    if (IsScannerConnected(source))
                    {
                        var scanner = new ScannerInfo
                        {
                            Id = source.Name,
                            Name = source.Name,
                            Type = ScannerType.TWAIN,
                            SupportsDuplex = CheckDuplexSupport(source) == NTwain.Data.BoolType.True ? true : false
                        };
                        scanners.Add(scanner);
                    }
                   
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to enumerate TWAIN scanners: {ex.Message}");
            }

            return scanners;
        }

        private bool IsScannerConnected(DataSource source)
        {
            try
            {
                var onlineCapability = source.Capabilities.CapDeviceOnline;
                var onlineResult = onlineCapability.GetCurrent();

                if (onlineResult == NTwain.Data.BoolType.True)
                {
                    System.Diagnostics.Debug.WriteLine($"Scanner {source.Name} is online");
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Scanner connection test failed for {source.Name}: {ex.Message}");
                return false;
            }
        }


        private NTwain.Data.BoolType CheckDuplexSupport(DataSource source)
        {

            source.Open();


            var cap = source.Capabilities.CapDuplexEnabled;
            var result = cap.GetCurrent();

            return result;

        }

        public async Task<List<ScanResult>> ScanAsync(ScannerInfo scannerInfo, ScanSettings settings)
        {
            return await Task.Run(() =>
            {
                var results = new List<ScanResult>();

                try
                {
                    if (_session == null || _session.State < 3)
                    {
                        results.Add(new ScanResult
                        {
                            Success = false,
                            ErrorMessage = "TWAIN session not available"
                        });
                        return results;
                    }

                    var source = _session.GetSources().FirstOrDefault(s => s.Name == scannerInfo.Id);
                    if (source == null)
                    {
                        results.Add(new ScanResult
                        {
                            Success = false,
                            ErrorMessage = "Scanner source not found"
                        });
                        return results;
                    }

                    lock (_lockObject)
                    {
                        source.Open();
                        source.Capabilities.CapDuplexEnabled?.SetValue(NTwain.Data.BoolType.True);
                        source.Capabilities.CapFeederEnabled?.SetValue(NTwain.Data.BoolType.True);
                        source.Capabilities.CapAutoFeed?.SetValue(NTwain.Data.BoolType.True);
                        source.Capabilities.ICapXResolution?.SetValue(150);
                        source.Capabilities.ICapYResolution?.SetValue(150);

                        if (source == null)
                        {
                            results.Add(new ScanResult
                            {
                                Success = false,
                                ErrorMessage = "Failed to open scanner"
                            });
                            return results;
                        }

                        // Configure scanner settings
                        ConfigureScanner(source, settings, scannerInfo);

                        // Enable source and start scanning
                        var rc = source.Enable(SourceEnableMode.ShowUI, false, IntPtr.Zero);
                        if (rc != ReturnCode.Success)
                        {
                            results.Add(new ScanResult
                            {
                                Success = false,
                                ErrorMessage = $"Failed to enable scanner: {rc}"
                            });
                            return results;
                        }

                        // Wait for and process images
                        _session.DataTransferred += (sender, e) =>
                            {
                                try
                                {
                                    var fileName = FileHelper.GenerateFileName(settings.SavePath, settings.Format);

                                    using (var image = e.GetNativeImageStream())
                                    using (var bitmap = new Bitmap(image))
                                    {
                                        Directory.CreateDirectory(Path.GetDirectoryName(fileName)!);
                                        bitmap.Save(fileName, settings.Format);

                                        results.Add(new ScanResult
                                        {
                                            Success = true,
                                            FilePath = fileName
                                        });
                                    }
                                }
                                catch (Exception ex)
                                {
                                    results.Add(new ScanResult
                                    {
                                        Success = false,
                                        ErrorMessage = $"Failed to save image: {ex.Message}"
                                    });
                                }
                            };

                        // Process messages until scanning is complete
                        NTwain.Data.Message msg = new NTwain.Data.Message();

                        while (_session.State == 5) Thread.Sleep(10000);
                        source.Close();
                        _session.Close();
                    }
                }
                catch (Exception ex)
                {
                    results.Add(new ScanResult
                    {
                        Success = false,
                        ErrorMessage = $"TWAIN scan failed: {ex.Message}"
                    });
                }

                return results;
            });
        }

        private void ConfigureScanner(DataSource ds, ScanSettings settings, ScannerInfo scannerInfo)
        {
            try
            {
                // Set pixel type (color mode)
                var pixelType = settings.ColorMode switch
                {
                    ColorMode.BlackAndWhite => PixelType.BlackWhite,
                    ColorMode.Grayscale => PixelType.Gray,
                    ColorMode.Color => PixelType.RGB,
                    _ => PixelType.RGB
                };
                ds.Capabilities.ICapPixelType.SetValue(pixelType);

                // Set resolution
                ds.Capabilities.ICapXResolution.SetValue((float)settings.Resolution);
                ds.Capabilities.ICapYResolution.SetValue((float)settings.Resolution);

                // Set duplex if supported and requested
                if (settings.UseDuplex && (scannerInfo.SupportsDuplex))
                {
                    ds.Capabilities.CapDuplexEnabled.SetValue(NTwain.Data.BoolType.True);
                }

                // Disable UI if not requested
                if (!settings.ShowUI)
                {
                    ds.Capabilities.CapIndicators.SetValue(NTwain.Data.BoolType.False);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to configure scanner: {ex.Message}");
            }
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
                        System.Diagnostics.Debug.WriteLine($"Error disposing TWAIN session: {ex.Message}");
                    }
                }
                _disposed = true;
            }
        }
    }
}
