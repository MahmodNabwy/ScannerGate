using ScannerApp.Helpers;
using ScannerApp.Models;
using System.Drawing.Imaging;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Net.Http.Headers;
using Microsoft.AspNetCore.Http;
using NTwain.Data;
using NTwain;
using System.Net.Http.Json;
using WIA;
using Newtonsoft.Json;

namespace ScannerApp
{
    public partial class ScannerForm : Form
    {
        private ScannerManager scannerManager;
        private ComboBox cmbScanners;
        private Button btnScan;
        private Button btnRefresh;
        private CheckBox chkDuplex;
        private TextBox txtSavePath;
        private ListBox lstResults;
        private ComboBox cmbResolution;
        private ComboBox cmbColorMode;
        private ComboBox cmbFormat;
        private ProgressBar progressBar;

        //M:\Scans
        private readonly string _saveDir = @"C:\Scans";

        //http://192.168.1.15:5000/ssn/extract
        private readonly string _ocrIDEndpoint = "http://192.168.100.149:5000/ssn/extract";
        private readonly string _ocrPassportEndpoint = "http://192.168.100.149:5000/passport/extract";
        private readonly string _ocrDrivingLicenseEndpoint = "http://192.168.100.149:5000/license/extract";
        private readonly string _backendApi = "http://192.168.100.149:699/api/admin/TempPerson";
        private int _imageCount = 0;
        public ScannerForm()
        {
            InitializeComponent();
            RunKestrelServer();
            scannerManager = new ScannerManager();


        }
        public async Task<object> TriggerScanAsync(int type = 1) //National ID
        {
            try
            {
                var scanner = LoadScanners(type);
                //if (type == 1)
                //{
                //    scanner.SupportsDuplex = true;
                //}

                if (!Directory.Exists(_saveDir))
                {
                    Directory.CreateDirectory(_saveDir);
                }


                string sessionId = Guid.NewGuid().ToString();
                string frontPath = "", backPath = "", passportPath = "", drivingLicensePath = "";

                // Marshal UI control access to UI thread
                bool useDuplex = false;
                ScannerInfo selectedScanner = scanner;



                var settings = new ScanSettings
                {
                    SavePath = _saveDir,
                    UseDuplex = useDuplex,
                    ColorMode = ScannerApp.Models.ColorMode.Color,
                    Resolution = 600, //Increasing Resolution will increase Image Size 
                    Format = ImageFormat.Jpeg
                };

                var results = await scannerManager.ScanAsync(selectedScanner, settings);
                if (results != null && results.Count > 0 && type == 4)
                {
                    var successfulResults = results.Where(r => r.Success && !string.IsNullOrEmpty(r.FilePath)).ToList();
                    drivingLicensePath = successfulResults[0].FilePath;

                }
                if (type == 1 && results.Count() == 2)
                {
                    frontPath = results[1].FilePath;
                    backPath = results[0].FilePath;
                }


                using var client = new HttpClient();

                MultipartFormDataContent IDParams = new MultipartFormDataContent();
                MultipartFormDataContent PassPortParams = new MultipartFormDataContent();
                MultipartFormDataContent DrivingLicenseParams = new MultipartFormDataContent();

                switch (type)
                {
                    case 1://Natgional ID
                        {
                            IDParams = new MultipartFormDataContent
                        {
                            { new StreamContent(File.OpenRead(frontPath)), "front", Path.GetFileName(frontPath) },
                            { new StreamContent(File.OpenRead(backPath)), "back", Path.GetFileName(backPath) }
                        };
                        }
                        break;
                    case 2://Passport
                        {
                            PassPortParams = new MultipartFormDataContent

                        {
                            { new StreamContent(File.OpenRead(passportPath)), "passport", Path.GetFileName(passportPath) },
                        };
                        }
                        break;
                    case 3://Just For Test
                        {
                            IDParams = new MultipartFormDataContent
                        {
                            { new StreamContent(File.OpenRead(frontPath)), "front", Path.GetFileName(frontPath) },
                            { new StreamContent(File.OpenRead(backPath)), "back", Path.GetFileName(backPath) }
                        };
                        }
                        break;
                    case 4://Driving License
                        {
                            DrivingLicenseParams = new MultipartFormDataContent
                        {
                            { new StreamContent(File.OpenRead(drivingLicensePath)), "license", Path.GetFileName(drivingLicensePath) },
                        };
                        }
                        break;

                    default:
                        {
                            PassPortParams = new MultipartFormDataContent
                        {
                            { new StreamContent(File.OpenRead(passportPath)), "passport", Path.GetFileName(passportPath) },
                        };
                        }
                        break;
                }


                var endpoint = type switch
                {
                    1 => _ocrIDEndpoint,//Natgional ID Endpoint
                    2 => _ocrPassportEndpoint,//Passport Endpoint
                    3 => _ocrIDEndpoint,// We Add Type 3 Just For Test
                    4 => _ocrDrivingLicenseEndpoint,//Driving License Endpoint
                    _ => _ocrPassportEndpoint//Default Passport Endpoint

                };

                var ocrResult = type switch
                {
                    1 => await PostWithRetryAsync<OcrResponse>(() => client.PostAsync(endpoint, IDParams), 3),
                    2 => await PostWithRetryAsync<OcrResponse>(() => client.PostAsync(endpoint, PassPortParams), 3),
                    3 => await PostWithRetryAsync<OcrResponse>(() => client.PostAsync(endpoint, IDParams), 3),
                    4 => await PostWithRetryAsync<OcrResponse>(() => client.PostAsync(endpoint, DrivingLicenseParams), 3),
                    _ => await PostWithRetryAsync<OcrResponse>(() => client.PostAsync(endpoint, PassPortParams), 3),
                };

               if (ocrResult == null || ocrResult.Status != "success")
                {
                    return new
                    {
                        status = true,
                        data = Array.Empty<OcrResponse>()
                    };
                }

                var payload = new
                {
                    sessionId,
                    firstName = ocrResult.Data.FirstName,
                    lastName = ocrResult.Data.LastName,
                    fullName = ocrResult.Data.FullName,
                    nationalId = ocrResult.Data.NationalID,
                    nationality = ocrResult.Data.Nationality,
                    issueingCountry = ocrResult.Data.IssuingCountry,
                    passportNumber = ocrResult.Data.PassportNumber,
                    birthDate = ocrResult.Data.BirthDate,
                    gender = ocrResult.Data.Gender == "Male" ? 1 : 2,
                    address = ocrResult.Data.Address,
                    job = ocrResult.Data.Job,
                    governorate = ocrResult.Data.Governorate,
                    expiryDate = ocrResult.Data.ExpiryDate,
                    serial = ocrResult.Data.Serial,
                    FrontImgUrl = ocrResult.FrontPath,
                    BackImgUrl = ocrResult.BackPath,
                    passportPath = ocrResult.PassportPath,
                    trafficUnit = ocrResult.Data.unit,
                    raw = JsonConvert.SerializeObject(ocrResult.Data)
                };


                var response = await PostAsync(() => client.PostAsJsonAsync(_backendApi, payload), 3);
                if (response != null)
                {
                    var dto = await response.Content.ReadFromJsonAsync<TempPersonDto>();

                    ocrResult.TempPersonId = dto.Id;
                    ocrResult.ActiveStatus = dto.Status;
                    ocrResult.ActiveStatusNote = dto.StatusNote;

                }

                ocrResult.SessionId = sessionId;
                return ocrResult;
            }
            catch (Exception ex)
            {
                File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: {ex.Message} \n");
                return new
                {
                    status = false,
                    data = Array.Empty<OcrResponse>()
                };
            }
        }
        private void InitializeComponent()
        {
            try
            {
                this.Size = new Size(600, 500);
                this.Text = "Scanner Manager";
                this.components = new System.ComponentModel.Container();
                this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                this.ClientSize = new System.Drawing.Size(0, 0);
                this.Text = "ID Scanner";
                File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: Application started {Environment.NewLine}");
            }
            catch (Exception ex)
            {
                File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: failed to start application: {ex.Message}{Environment.NewLine}");
            }

            // Scanner selection
            //var lblScanner = new Label { Text = "Scanner:", Location = new Point(10, 15), Size = new Size(80, 20) };
            //cmbScanners = new ComboBox { Location = new Point(100, 12), Size = new Size(300, 25), DropDownStyle = ComboBoxStyle.DropDownList };

            //// Duplex option
            //chkDuplex = new CheckBox { Text = "Duplex (if supported)", Location = new Point(420, 15), Size = new Size(150, 20) };

            //// Save path
            //var lblPath = new Label { Text = "Save Path:", Location = new Point(10, 50), Size = new Size(80, 20) };
            //txtSavePath = new TextBox { Location = new Point(100, 47), Size = new Size(300, 25), Text = @"C:\ScannedImages\" };
            //var btnBrowse = new Button { Text = "Browse", Location = new Point(420, 46), Size = new Size(80, 27) };
            //btnBrowse.Click += BtnBrowse_Click;

            //// Scan button
            //btnScan = new Button { Text = "Scan", Location = new Point(100, 85), Size = new Size(100, 35) };
            //btnScan.Click += BtnScan_Click;

            //// Results
            //var lblResults = new Label { Text = "Scan Results:", Location = new Point(10, 135), Size = new Size(100, 20) };
            //lstResults = new ListBox { Location = new Point(10, 160), Size = new Size(560, 300) };

            //this.Controls.AddRange(new Control[] { lblScanner, cmbScanners, chkDuplex, lblPath, txtSavePath, btnBrowse, btnScan, lblResults, lstResults });
        }
        private ScannerInfo? LoadScanners(int type)
        {
            var scanners = scannerManager.GetAvailableScanners();
            if (scanners.Count > 0)
            {
                File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: failed:scanners.Count = 0 {Environment.NewLine}");

            }
            //return scanners.FirstOrDefault(c => c.Name == "EPSON DS-530");
            return scanners
                    .Where(c => !c.Name.Contains("Vue"))
                    .OrderBy(c => c.Type == ScannerType.TWAIN ? 0 : 1)
                    .FirstOrDefault();

            //cmbScanners.DataSource = scanners;
            //cmbScanners.DisplayMember = "Name";
            //cmbScanners.ValueMember = "Id";
            ////// Set the selected scanner to be the first item

        }
        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtSavePath.Text = dialog.SelectedPath;
                }
            }
        }
        private async void BtnScan_Click(object sender, EventArgs e)
        {
            if (cmbScanners.SelectedItem == null)
            {
                MessageBox.Show("Please select a scanner.");
                return;
            }
            var scanners = scannerManager.GetAvailableScanners();

            var selectedScanner = (ScannerInfo)scanners.FirstOrDefault();
            //btnScan.Enabled = false;
            lstResults.Items.Clear();

            try
            {
                Directory.CreateDirectory(txtSavePath.Text);

                var settings = new ScanSettings
                {
                    SavePath = txtSavePath.Text,
                    UseDuplex = false,
                    ColorMode = ScannerApp.Models.ColorMode.Color,
                    Resolution = 300,
                    Format = ImageFormat.Jpeg
                };

                lstResults.Items.Add($"Starting scan with {selectedScanner.Name}...");
                Application.DoEvents();

                var results = await scannerManager.ScanAsync(selectedScanner, settings);

                foreach (var result in results)
                {
                    lstResults.Items.Add($"Saved: {result}");
                }

                lstResults.Items.Add($"Scan completed. {results.Count} image(s) saved.");
            }
            catch (Exception ex)
            {
                lstResults.Items.Add($"Error: {ex.Message}");
                MessageBox.Show($"Scanning failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnScan.Enabled = true;
            }
        }

        private void RunKestrelServer()
        {
            Task.Run(() =>
            {
                var builder = WebApplication.CreateBuilder();
                builder.Services.AddCors(options =>
                {
                    options.AddPolicy("CORSPolicy", builder =>
                        builder.WithHeaders(HeaderNames.ContentType, "x-custom-header")
                            .WithHeaders(HeaderNames.Accept, "*/*")
                            .WithHeaders(HeaderNames.AcceptEncoding, "gzip, deflate, br")
                            .WithHeaders(HeaderNames.Connection, "keep-alive").AllowAnyMethod().AllowAnyHeader().AllowCredentials()
                            .SetIsOriginAllowed((hosts) => true));
                });
                builder.Services.AddSingleton(this);
                builder.Services.AddControllers();
                var app = builder.Build();

                app.UseCors("CORSPolicy");

                app.MapPost("/scan", async (HttpContext context, ScannerForm form) =>
                {
                    var formData = await context.Request.ReadFromJsonAsync<ScanRequest>();
                    int type = formData?.Type ?? 1; //Natgional ID
                    return await form.TriggerScanAsync(type);
                    //return Results.Json(new { status = "Scan complete", documentType = type == 1 ? "National ID" : "Passport" });
                });

                //app.MapPost("/mock-ocr", async () =>
                //{Dh#30@Ldc@2024
                //    return Results.Json(new
                //    {
                //        fullName = "Test User",
                //        nationalId = "12345678901234",
                //        dob = "1990-01-01",
                //        address = "123 Main St"
                //    });
                //});

                //41.196.0.189
                //app.Run("http://192.168.15.120:5005");
                //app.Run("http://0.0.0.0:5005");
                app.Run("http://localhost:5005");
                //app.Run("http://192.168.1.15:5005");


            });
        }

        private async Task<HttpResponseMessage?> PostAsync(
        Func<Task<HttpResponseMessage>> action, int maxRetries)
        {
            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    var response = await action();
                    if (response.IsSuccessStatusCode)
                        return response; // Return raw response, no DTO mapping inside
                }
                catch (Exception ex)
                {
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"),
                        $"{DateTime.Now}: Retry {attempt + 1} failed: {ex.Message}{Environment.NewLine}");
                }

                await Task.Delay(1000 * (int)Math.Pow(2, attempt));
            }

            return null;
        }

        private async Task<T?> PostWithRetryAsync<T>(Func<Task<HttpResponseMessage>> action, int maxRetries)
        {
            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    var response = await action();
                    if (response.IsSuccessStatusCode)
                        return await response.Content.ReadFromJsonAsync<T>();
                }
                catch (Exception ex)
                {
                    File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: Retry {attempt + 1} failed: {ex.Message}{Environment.NewLine}");
                }

                await Task.Delay(1000 * (int)Math.Pow(2, attempt)); // exponential backoff
            }

            return default;
        }
    }


    public class ScannerManager
    {
        private readonly string _saveDir = @"C:\Scans";
        private readonly WiaManager wiaManager;
        private readonly TwainManager twainManager;

        public ScannerManager()
        {
            wiaManager = new WiaManager();
            twainManager = new TwainManager();
        }

        public List<ScannerInfo> GetAvailableScanners()
        {
            var scanners = new List<ScannerInfo>();

            // Get WIA scanners
            try
            {
                scanners.AddRange(wiaManager.GetScanners());
            }
            catch (Exception ex)
            {
                File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: failed to load WIA Scanners: {ex.Message}{Environment.NewLine}");

                System.Diagnostics.Debug.WriteLine($"WIA scanner detection failed: {ex.Message}");
            }

            // Get TWAIN scanners
            try
            {
                scanners.AddRange(twainManager.GetScanners());
            }
            catch (Exception ex)
            {
                File.AppendAllText(Path.Combine(_saveDir, "scanner.log"), $"{DateTime.Now}: failed to load TWAIN Scanners: {ex.Message}{Environment.NewLine}");
                System.Diagnostics.Debug.WriteLine($"TWAIN scanner detection failed: {ex.Message}");
            }

            return scanners.DistinctBy(s => s.Id).ToList();
        }

        public async Task<List<ScanResult>> ScanAsync(ScannerInfo scanner, ScanSettings settings)
        {
            //return new List<string>();
            return scanner.Type switch
            {
                ScannerType.WIA => await wiaManager.ScanAsync(scanner, settings),
                ScannerType.TWAIN => await twainManager.ScanAsync(scanner, settings),
                _ => throw new NotSupportedException($"Scanner type {scanner.Type} is not supported.")
            };
        }
    }
}
