using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;

namespace WebPConverter
{
    public class Converter
    {
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                // Gdy aplikacja jest uruchamiana bez argumentów, uruchom instalator
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new InstallerForm());
                return;
            }

            // Gdy aplikacja jest uruchamiana z argumentami, wykonaj konwersję
            try
            {
                string filePath = args[0];

                if (!File.Exists(filePath))
                {
                    MessageBox.Show($"Plik nie istnieje: {filePath}", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!IsImageFile(filePath))
                {
                    MessageBox.Show("Plik nie jest obsługiwanym formatem graficznym (JPG lub PNG).",
                        "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ConvertToWebP(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Wystąpił błąd podczas konwersji:\n\n{ex.Message}",
                    "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static bool IsImageFile(string filePath)
        {
            string extension = Path.GetExtension(filePath).ToLower();
            return extension == ".jpg" || extension == ".jpeg" || extension == ".png";
        }

        private static void ConvertToWebP(string filePath)
        {
            // Ścieżka do pliku wyjściowego
            string outputPath = Path.ChangeExtension(filePath, ".webp");

            // Sprawdź, czy plik już istnieje
            if (File.Exists(outputPath))
            {
                var result = MessageBox.Show($"Plik {Path.GetFileName(outputPath)} już istnieje. Czy chcesz go nadpisać?",
                    "Plik istnieje", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.No)
                {
                    return;
                }
            }

            // Użyj ImageMagick do konwersji
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "magick.exe",
                    Arguments = $"\"{filePath}\" -quality 90 \"{outputPath}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                }
            };

            process.Start();
            process.WaitForExit();

            if (process.ExitCode != 0 || !File.Exists(outputPath))
            {
                string error = process.StandardError.ReadToEnd();
                throw new Exception($"Konwersja nie powiodła się. Kod wyjścia: {process.ExitCode}\n{error}");
            }

            // Wyświetl potwierdzenie tylko jeśli konwersja zakończyła się sukcesem
            MessageBox.Show($"Plik został pomyślnie przekonwertowany do WebP:\n{outputPath}",
                "Konwersja zakończona", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void InstallContextMenu(bool install = true)
        {
            string executablePath = Application.ExecutablePath;

            try
            {
                if (install)
                {
                    // Dodaj wpisy do rejestru dla JPG
                    using (var key = Registry.ClassesRoot.CreateSubKey(@"SystemFileAssociations\.jpg\shell\ConvertToWebP"))
                    {
                        key.SetValue("", "Konwertuj do WebP");
                        key.SetValue("Icon", "shell32.dll,43");
                    }

                    using (var key = Registry.ClassesRoot.CreateSubKey(@"SystemFileAssociations\.jpg\shell\ConvertToWebP\command"))
                    {
                        key.SetValue("", $"\"{executablePath}\" \"%1\"");
                    }

                    // Dodaj wpisy do rejestru dla JPEG
                    using (var key = Registry.ClassesRoot.CreateSubKey(@"SystemFileAssociations\.jpeg\shell\ConvertToWebP"))
                    {
                        key.SetValue("", "Konwertuj do WebP");
                        key.SetValue("Icon", "shell32.dll,43");
                    }

                    using (var key = Registry.ClassesRoot.CreateSubKey(@"SystemFileAssociations\.jpeg\shell\ConvertToWebP\command"))
                    {
                        key.SetValue("", $"\"{executablePath}\" \"%1\"");
                    }

                    // Dodaj wpisy do rejestru dla PNG
                    using (var key = Registry.ClassesRoot.CreateSubKey(@"SystemFileAssociations\.png\shell\ConvertToWebP"))
                    {
                        key.SetValue("", "Konwertuj do WebP");
                        key.SetValue("Icon", "shell32.dll,43");
                    }

                    using (var key = Registry.ClassesRoot.CreateSubKey(@"SystemFileAssociations\.png\shell\ConvertToWebP\command"))
                    {
                        key.SetValue("", $"\"{executablePath}\" \"%1\"");
                    }
                }
                else
                {
                    // Usuń wpisy z rejestru
                    Registry.ClassesRoot.DeleteSubKeyTree(@"SystemFileAssociations\.jpg\shell\ConvertToWebP", false);
                    Registry.ClassesRoot.DeleteSubKeyTree(@"SystemFileAssociations\.jpeg\shell\ConvertToWebP", false);
                    Registry.ClassesRoot.DeleteSubKeyTree(@"SystemFileAssociations\.png\shell\ConvertToWebP", false);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Błąd podczas modyfikowania rejestru: {ex.Message}", ex);
            }
        }
    }
}