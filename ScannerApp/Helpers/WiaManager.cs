using ScannerApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using WIA;
namespace ScannerApp.Helpers
{
    public class WiaManager
    {
        private readonly string _saveDir = @"C:\Scans";
        public List<ScannerInfo> GetScanners()
        {
            var scanners = new List<ScannerInfo>();

            try
            {
                var deviceManager = new DeviceManager();

                foreach (DeviceInfo deviceInfo in deviceManager.DeviceInfos)
                {
                    if (deviceInfo.Type == WiaDeviceType.ScannerDeviceType)
                    {
                        try
                        {
                            // Attempt to connect to the device to verify it's online.
                            var device = deviceInfo.Connect();

                            var scanner = new ScannerInfo
                            {
                                Id = deviceInfo.DeviceID,
                                Name = GetDeviceName(deviceInfo),
                                Type = ScannerType.WIA,
                                SupportsDuplex = CheckDuplexSupport(deviceInfo)
                            };
                            scanners.Add(scanner);

                            // Release the COM object as it's no longer needed here.
                            if (device != null)
                            {
                                Marshal.ReleaseComObject(device);
                            }
                        }
                        catch (COMException ex)
                        {
                            // A COMException is thrown if the device is not connected.
                            // We can log this for debugging but otherwise ignore it.
                            System.Diagnostics.Debug.WriteLine($"Scanner '{GetDeviceName(deviceInfo)}' is offline or inaccessible: {ex.Message}");
                        }
                    }
                }
            }
            catch (COMException ex)
            {
                throw new Exception($"Failed to enumerate WIA scanners: {ex.Message}", ex);
            }

            return scanners;
        }

        private string GetDeviceName(DeviceInfo deviceInfo)
        {
            try
            {
                return deviceInfo.Properties["Name"].get_Value()?.ToString() ?? "Unknown Scanner";
            }
            catch
            {
                return "Unknown Scanner";
            }
        }

        private bool CheckDuplexSupport(DeviceInfo deviceInfo)
        {
            try
            {
                var device = deviceInfo.Connect();
                
                // Common duplex-related property IDs
                int[] duplexPropertyIds = { 
                    6028, // WIA_DPS_DOCUMENT_HANDLING_SELECT
                    6027, // WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES
                    3078  // WIA_IPA_DOCUMENT_HANDLING_SELECT
                };
                
                foreach (Item item in device.Items)
                {
                    foreach (Property prop in item.Properties)
                    {
                        if (duplexPropertyIds.Contains(prop.PropertyID))
                        {
                            System.Diagnostics.Debug.WriteLine($"Found duplex-related property: {prop.PropertyID}");
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error checking duplex support: {ex.Message}");
            }

            return false;
        }

        public async Task<List<ScanResult>> ScanAsync(ScannerInfo scannerInfo, ScanSettings settings)
        {
            return await Task.Run(() =>
            {
                var results = new List<ScanResult>();

                try
                {
                    var deviceManager = new DeviceManager();
                    var deviceInfo = deviceManager.DeviceInfos.Cast<DeviceInfo>()
                        .FirstOrDefault(d => d.DeviceID == scannerInfo.Id);

                    if (deviceInfo == null)
                    {
                        results.Add(new ScanResult
                        {
                            Success = false,
                            ErrorMessage = "Scanner not found"
                        });
                        File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: failed: Scanner not found {Environment.NewLine}");
                        return results;
                    }

                    var device = deviceInfo.Connect();
                    if (device.Items.Count == 0)
                    {
                        results.Add(new ScanResult
                        {
                            Success = false,
                            ErrorMessage = "No scan items available"
                        });
                        File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: failed: No Scan items available {Environment.NewLine}");
                        return results;

                    }

                    var item = device.Items[1];

                    // Set scan properties
                    //SetScanProperty(item, 6146, (int)settings.ColorMode); // WIA_IPA_DATATYPE
                    SetScanProperty(item, 6147, settings.Resolution); // WIA_IPS_XRES
                    SetScanProperty(item, 6148, settings.Resolution); // WIA_IPS_YRES

                    // Handle duplex if supported and requested
                    //if (settings.UseDuplex && scannerInfo.SupportsDuplex)
                    if (1 == 1)
                    {
                        SetScanProperty(item, 6028, 1); // Enable duplex
                    }

                    // Determine format
                    string formatGuid = settings.Format.Guid.ToString("B").ToUpper();

                    // Perform scan
                    var imageFile = (ImageFile)item.Transfer(formatGuid);

                    var fileName = FileHelper.GenerateFileName(settings.SavePath, settings.Format);
                    imageFile.SaveFile(fileName);

                    results.Add(new ScanResult
                    {
                        Success = true,
                        FilePath = fileName
                    });

                    // Handle back page for duplex
                    //if (settings.UseDuplex && scannerInfo.SupportsDuplex)
                    if (1 == 1)
                    {
                        try
                        {
                            var backImageFile = (ImageFile)item.Transfer(formatGuid);
                            var backFileName = FileHelper.GenerateFileName(settings.SavePath, settings.Format);
                            backImageFile.SaveFile(backFileName);

                            results.Add(new ScanResult
                            {
                                Success = true,
                                FilePath = backFileName
                            });
                        }
                        catch (Exception ex)
                        {
                            File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: failed: Dublex Error {Environment.NewLine}");
                            System.Diagnostics.Debug.WriteLine($"No back page or error: {ex.Message}");
                        }
                    }
                }
                catch (COMException ex)
                {
                    results.Add(new ScanResult
                    {
                        Success = false,
                        ErrorMessage = $"WIA scan failed: {ex.Message}"
                    });
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: failed: WIA scan failed: {ex.Message} {Environment.NewLine}");
                }
                catch (Exception ex)
                {
                    results.Add(new ScanResult
                    {
                        Success = false,
                        ErrorMessage = $"Scan error: {ex.Message}"
                    });
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: failed: Scan error: {ex.Message} {Environment.NewLine}");

                }

                return results;
            });
        }

        private void SetScanProperty(Item item, int propertyId, object value)
        {
            try
            {
                foreach (Property prop in item.Properties)
                {
                    if (prop.PropertyID == propertyId)
                    {
                        prop.set_Value(value);
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: failed: Failed to set property {propertyId}: {ex.Message} {Environment.NewLine}");

                System.Diagnostics.Debug.WriteLine($"Failed to set property {propertyId}: {ex.Message}");
            }
        }

    }
}
