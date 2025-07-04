using System.Diagnostics;
using System.Net;
using Microsoft.Win32;
using conv2WebP4Win;

namespace WebPConverter
{
    public partial class InstallerForm : Form
    {
        private const string AppName = "WebP Converter";
        private const string AppVersion = "1.1.0";
        private string installPath;
        private string converterPath;
        private readonly string tempPath = Path.Combine(Path.GetTempPath(), "WebPConverterTemp");

        public InstallerForm()
        {
            // Wymuś angielski dla wszystkich języków poza polskim
            var culture = System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName;
            if (culture != "pl")
            {
                Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en");
            }
            InitializeComponent();
            installPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "WebPConverter");
            converterPath = Path.Combine(installPath, "WebPConverter.exe");
            this.Text = $"{AppName} Installer v{AppVersion} by devLNX Studio";
        }

        private void InitializeComponent()
        {
            this.Width = 500;
            this.Height = 350;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;

            var label = new Label
            {
                //Text = "Konwerter JPG/PNG do WebP - Instalator",
                Text = Strings.labelTitle,
                Font = new System.Drawing.Font("Segoe UI", 14, System.Drawing.FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 50,
                AutoSize = false
            };

            var descriptionLabel = new Label
            {
                Text = Strings.labelSTitle,
                TextAlign = System.Drawing.ContentAlignment.TopLeft,
                Location = new System.Drawing.Point(20, 60),
                Size = new System.Drawing.Size(460, 80)
            };

            var progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(20, 230),
                Size = new System.Drawing.Size(460, 25),
                Minimum = 0,
                Maximum = 100,
                Value = 0,
                Name = "progressBar"
            };

            var statusLabel = new Label
            {
                Text =Strings.labelReady,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Location = new System.Drawing.Point(20, 200),
                Size = new System.Drawing.Size(460, 25),
                Name = "statusLabel"
            };

            var buttonSpacing = 20;
            var buttonWidth = 100;
            var buttonHeight = 30;

            int totalWidth = buttonWidth * 2 + buttonSpacing;
            int startX = (this.ClientSize.Width - totalWidth) / 2;

            var installButton = new Button
            {
                Text = Strings.btnInstall,
                Size = new Size(buttonWidth, buttonHeight),
                Location = new Point(startX, 270),
                Name = "installButton"
            };

            var cancelButton = new Button
            {
                Text = Strings.btnCancel,
                Size = new Size(buttonWidth, buttonHeight),
                Location = new Point(startX + buttonWidth + buttonSpacing, 270),
                Name = "cancelButton"
            };

            installButton.Click += InstallButton_Click;
            cancelButton.Click += (s, e) => Close();

            this.Controls.Add(label);
            this.Controls.Add(descriptionLabel);
            this.Controls.Add(progressBar);
            progressBar.Left = (this.ClientSize.Width - progressBar.Width) / 2;
            this.Controls.Add(statusLabel);
            this.Controls.Add(installButton);
            this.Controls.Add(cancelButton);
        }

        private void InstallButton_Click(object? sender, EventArgs e)
        {
            // Deaktywuj przycisk instalacji
            var installButton = (Button)Controls.Find("installButton", true)[0];
            installButton.Enabled = false;

            // Rozpocznij instalację w osobnym wątku
            System.Threading.Tasks.Task.Run(() => Install());
        }

        private void Install()
        {
            try
            {
                var progressBar = (ProgressBar)Controls.Find("progressBar", true)[0];
                var statusLabel = (Label)Controls.Find("statusLabel", true)[0];

                // Sprawdź czy aplikacja jest uruchomiona jako administrator
                if (!IsAdministrator())
                {
                    MessageBox.Show(Strings.msgAdmin,
                        Strings.genError, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Invoke(new Action(() => Close()));
                    return;
                }

                // Aktualizacja statusu
                UpdateStatus(Strings.statUpd, 10);

                // Utwórz katalog instalacyjny
                if (!Directory.Exists(installPath))
                {
                    Directory.CreateDirectory(installPath);
                }

                // Utwórz katalog tymczasowy
                if (!Directory.Exists(tempPath))
                {
                    Directory.CreateDirectory(tempPath);
                }

                // Sprawdź czy ImageMagick jest już zainstalowany
                UpdateStatus(Strings.statUpdIM, 20);
                bool imagemagickInstalled = IsImageMagickInstalled();

                if (!imagemagickInstalled)
                {
                    // Pobierz i zainstaluj ImageMagick
                    UpdateStatus(Strings.statUpdIM2, 30);
                    string imagemagickInstallerPath = DownloadImageMagick();

                    UpdateStatus(Strings.statUpdIM3, 40);
                    InstallImageMagick(imagemagickInstallerPath);
                }
                else
                {
                    UpdateStatus(Strings.statUpdIM4, 40);
                }

                // Utwórz plik konwertera
                UpdateStatus(Strings.statUpd1, 60);
                CreateConverterExe();

                // Dodaj wpisy do rejestru
                UpdateStatus(Strings.statUpd2, 80);
                AddRegistryEntries();

                // Zakończenie instalacji
                UpdateStatus(Strings.statUpd3, 100);
                MessageBox.Show(Strings.msgSuccess,
                    Strings.genSuccess, MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Wyczyść pliki tymczasowe
                try
                {
                    if (Directory.Exists(tempPath))
                    {
                        Directory.Delete(tempPath, true);
                    }
                }
                catch (Exception) { /* Ignoruj błędy przy czyszczeniu */ }

                // Zamknij aplikację
                Invoke(new Action(() => Close()));
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(Strings.genErrorMsg, ex.Message), Strings.genError, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Invoke(new Action(() =>
                {
                    var installButton = (Button)Controls.Find("installButton", true)[0];
                    installButton.Enabled = true;
                    UpdateStatus(Strings.genFailed, 0);
                }));
            }
        }

        private void UpdateStatus(string status, int progress)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => UpdateStatus(status, progress)));
                return;
            }

            var progressBar = (ProgressBar)Controls.Find("progressBar", true)[0];
            var statusLabel = (Label)Controls.Find("statusLabel", true)[0];

            statusLabel.Text = status;
            progressBar.Value = progress;
        }

        private bool IsAdministrator()
        {
            var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var principal = new System.Security.Principal.WindowsPrincipal(identity);
            return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
        }

        private bool IsImageMagickInstalled()
        {
            try
            {
                // Sprawdź czy magick.exe jest w ścieżce systemowej
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "where",
                        Arguments = "magick.exe",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        CreateNoWindow = true
                    }
                };

                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();

                return process.ExitCode == 0 && !string.IsNullOrEmpty(output);
            }
            catch
            {
                return false;
            }
        }

        private string DownloadImageMagick()
        {
            // Pobierz instalator ImageMagick
            string installerPath = Path.Combine(tempPath, "ImageMagick-Installer.exe");
            string downloadUrl = "https://imagemagick.org/archive/binaries/ImageMagick-7.1.1-20-Q16-HDRI-x64-dll.exe";

            using (var client = new WebClient())
            {
                client.DownloadFile(downloadUrl, installerPath);
            }

            return installerPath;
        }

        private void InstallImageMagick(string installerPath)
        {
            // Uruchom instalator ImageMagick z cichą instalacją
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = installerPath,
                    Arguments = "/SILENT /NORESTART /NOICONS",
                    UseShellExecute = true,
                    Verb = "runas"
                }
            };

            process.Start();
            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                throw new Exception(Strings.genIMError);
            }
        }

        private void CreateConverterExe()
        {
            // W rzeczywistej aplikacji tutaj należałoby utworzyć lub skopiować plik .exe
            // Na potrzeby tego przykładu tworzymy skrypt PowerShell, który będzie wywoływany

            string converterScriptPath = Path.Combine(installPath, "convert-to-webp.ps1");
            string converterScript = @"
param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath
)

# Sciezka do pliku wyjsciowego
$outputPath = [System.IO.Path]::ChangeExtension($FilePath, "".webp"")

# Konwersja do WebP z uzyciem ImageMagick
& magick.exe ""$FilePath"" -quality 90 ""$outputPath""

# Sprawdz, czy konwersja sie powiodla
if (Test-Path -Path $outputPath) {
    Write-Host ""Konwersja zakonczona sukcesem: $outputPath"" -ForegroundColor Green
} else {
    Write-Host ""Konwersja nie powiodla sie"" -ForegroundColor Red
}
";

            File.WriteAllText(converterScriptPath, converterScript);

            // W rzeczywistej aplikacji tutaj utworzylibyśmy plik .exe
            // Na potrzeby tego przykładu używamy skryptu PowerShell
            converterPath = converterScriptPath;
        }

        private void AddRegistryEntries()
        {
            // Dodaj wpisy do rejestru dla JPG
            using (var key = Registry.ClassesRoot.CreateSubKey(@"SystemFileAssociations\.jpg\shell\ConvertToWebP"))
            {
                key.SetValue("", Strings.regEntry);
                key.SetValue("Icon", "shell32.dll,43");
            }

            using (var key = Registry.ClassesRoot.CreateSubKey(@"SystemFileAssociations\.jpg\shell\ConvertToWebP\command"))
            {
                key.SetValue("", $"powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File \"{converterPath}\" -FilePath \"%1\"");

            }

            // Dodaj wpisy do rejestru dla JPEG
            using (var key = Registry.ClassesRoot.CreateSubKey(@"SystemFileAssociations\.jpeg\shell\ConvertToWebP"))
            {
                key.SetValue("", Strings.regEntry);
                key.SetValue("Icon", "shell32.dll,43");
            }

            using (var key = Registry.ClassesRoot.CreateSubKey(@"SystemFileAssociations\.jpeg\shell\ConvertToWebP\command"))
            {
                key.SetValue("", $"powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File \"{converterPath}\" -FilePath \"%1\"");

            }


            // Dodaj wpisy do rejestru dla PNG
            using (var key = Registry.ClassesRoot.CreateSubKey(@"SystemFileAssociations\.png\shell\ConvertToWebP"))
            {
                key.SetValue("",Strings.regEntry);
                key.SetValue("Icon", "shell32.dll,43");
            }

            using (var key = Registry.ClassesRoot.CreateSubKey(@"SystemFileAssociations\.png\shell\ConvertToWebP\command"))
            {
                key.SetValue("", $"powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File \"{converterPath}\" -FilePath \"%1\"");

            }
        }
    }
}