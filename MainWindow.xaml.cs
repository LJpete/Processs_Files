using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using WinRT.Interop;
using System.Runtime.InteropServices;

namespace App1
{
    public sealed partial class MainWindow : Window
    {
        private int numberOfPlates;

        private static readonly string CliExePath = LoadSysLocation();

        private static string LoadSysLocation()
        {
            var exeDir = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            var locationFile = Path.Combine(exeDir, "system.location");

            if (!File.Exists(locationFile))
            {
                throw new FileNotFoundException($"Missing {locationFile}");
            }

            string relativePath;
            using (var reader = new StreamReader(locationFile))
            {
                reader.ReadLine(); // skip first line
                relativePath = reader.ReadLine()?.Trim();
            }

            if (string.IsNullOrEmpty(relativePath))
            {
                throw new InvalidOperationException(
                    $"system.location [myassays dir location]is empty or invalid: {locationFile}");
            }

            var programFilesX86 =
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

            var exePath = Path.Combine(programFilesX86, relativePath);

            if (!File.Exists(exePath))
            {
                throw new FileNotFoundException($"Executable not found: {exePath}");
            }

            return exePath;
        }


        private static readonly string[] ArgumentTemplate =
            LoadArgumentTemplate();

        private static string[] LoadArgumentTemplate()
        {
            var exeDir = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            var templatePath = Path.Combine(exeDir, "Arguments.template");

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException(
                    $"argument template not found: {templatePath}");
            }

            return File.ReadAllLines(templatePath)
                       .Where(line => !string.IsNullOrWhiteSpace(line))
                       .ToArray();
        }

        private string protocol_Plate;

        public MainWindow()
        {
            this.InitializeComponent();

            ExtendsContentIntoTitleBar = true;

            var hwnd = WindowNative.GetWindowHandle(this);
            var windowId = Win32Interop.GetWindowIdFromWindow(hwnd);
            var appWindow = AppWindow.GetFromWindowId(windowId);

            appWindow.SetPresenter(AppWindowPresenterKind.Overlapped);

            // Get DPI for the window
            var dpi = GetDpiForWindow(hwnd);
            double scaling = dpi / 96.0;

            // Desired size in DIPs
            int widthDIPs = 490;
            int heightDIPs = 370;

            // Convert DIPs to physical pixels
            int widthPx = (int)(widthDIPs * scaling);
            int heightPx = (int)(heightDIPs * scaling);

            this.AppWindow.MoveAndResize(new Windows.Graphics.RectInt32(800, 700, widthPx, heightPx));
        }

        // Add this P/Invoke
        [DllImport("user32.dll")]
        static extern uint GetDpiForWindow(IntPtr hWnd);

        public void PlateCount(object sender, RoutedEventArgs e)
        {
            if (sender is MenuFlyoutItem item &&
                int.TryParse(item.Tag?.ToString(), out int plates))
            {
                numberOfPlates = plates;

                // Optional: update button text (control assumed defined in XAML)
                nPlatesDropDownButton.Content = $"{plates} Chips read";
            }
        }

        private string _selectedPlateText = string.Empty;

        private void PlateSelected(object sender, RoutedEventArgs e)
        {
            if (sender is MenuFlyoutItem item)
            {
                protocol_Plate = item.Tag?.ToString() ?? string.Empty;
                PlatesDropDownButton.Content = item.Text;

                _selectedPlateText = item.Text?.Trim();
            }
        }

        public ListView GetRecentFilesList(string SysOutput)
        {
            var takeCount = Math.Max(1, numberOfPlates);
            var files = Directory.GetFiles(SysOutput, "*.xlsx")
                         .Select(f => new FileInfo(f))
                         .OrderByDescending(f => f.LastWriteTime)
                         .Take(takeCount)
                         .Select(f => $"{f.Name}  ({f.LastWriteTime:g})")
                         .ToList();

            return new ListView
            {
                ItemsSource = files,
                IsHitTestVisible = false, // display-only
                Margin = new Thickness(0, 8, 0, 0)
            };
        }

        private async void OnSubmitClicked(object sender, RoutedEventArgs e)
        {
            string SysOutput;
            try
            {
                SysOutput = LoadSystemOutput();
            }
            catch (Exception ex)
            {
                var err = new ContentDialog
                {
                    Title = "System Output Error",
                    Content = ex.Message,
                    CloseButtonText = "OK",
                    XamlRoot = Content.XamlRoot
                };
                _ = err.ShowAsync();
                return;
            }

            var hBarcode = MyTextBox?.Text?.Trim() ?? string.Empty;
            var invalidChars = Path.GetInvalidFileNameChars();
            var safeBarcode = string.Concat(hBarcode.Where(c => !invalidChars.Contains(c)));

            var plate = PlatesDropDownButton?.Content?.ToString()?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(plate) || string.IsNullOrWhiteSpace(hBarcode))
            {
                var err = new ContentDialog
                {
                    Title = "Missing input",
                    Content = "Please select a plate and enter a barcode before launching.",
                    CloseButtonText = "OK",
                    XamlRoot = Content.XamlRoot
                };
                _ = err.ShowAsync();
                return;
            }

            var recentFilesList = GetRecentFilesList(SysOutput);

            var contentPanel = new StackPanel
            {
                Spacing = 8,
                Children =
                {
                    new TextBlock
                    {
                        Text = "Most recent files:",
                        FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
                    },
                    recentFilesList
                }
            };

            var dialog = new ContentDialog
            {
                Title = "Files to Rename",
                Content = contentPanel,
                CloseButtonText = "Not My Plate",
                PrimaryButtonText = "Process sample files",
                XamlRoot = Content.XamlRoot
            };

            var result = await dialog.ShowAsync();
            if (result == ContentDialogResult.Primary)
            {
                var maxPlates = Math.Max(1, numberOfPlates);

                var filesToRename = Directory.GetFiles(SysOutput)
                    .Select(p => new FileInfo(p))
                    .OrderByDescending(f => f.LastWriteTime)
                    .Take(maxPlates)
                    .OrderBy(f => f.LastWriteTime)
                    .ToArray();

                if (filesToRename.Length == 0)
                {
                    var noneDlg = new ContentDialog
                    {
                        Title = "No files found",
                        Content = $"No files were found in {SysOutput}.",
                        CloseButtonText = "OK",
                        XamlRoot = Content.XamlRoot
                    };
                    _ = noneDlg.ShowAsync();
                    return;
                }

                var renameResults = new List<(string Original, string Destination, string Error)>();

                await Task.Run(() =>
                {
                    for (var idx = 0; idx < filesToRename.Length; idx++)
                    {
                        var fi = filesToRename[idx];
                        var plateIndex = idx + 1;
                        var extension = fi.Extension ?? string.Empty;
                        var desiredName = $"{safeBarcode}.{plateIndex}{extension}";
                        var destPath = Path.Combine(SysOutput, desiredName);

                        if (File.Exists(destPath))
                        {
                            var suffix = 1;
                            string candidate;
                            do
                            {
                                candidate = Path.Combine(SysOutput, $"{safeBarcode}.{plateIndex}({suffix}){extension}");
                                suffix++;
                            } while (File.Exists(candidate));
                            destPath = candidate;
                        }

                        try
                        {
                            File.Move(fi.FullName, destPath);
                            renameResults.Add((fi.Name, Path.GetFileName(destPath), null));
                        }
                        catch (Exception ex)
                        {
                            renameResults.Add((fi.Name, Path.GetFileName(destPath), ex.Message));
                        }
                    }
                });

                var successes = renameResults.Where(r => r.Error == null).ToList();
                var failures = renameResults.Where(r => r.Error != null).ToList();

                if (failures.Count == 0)
                {
                    var summary = new StackPanel { Spacing = 6 };
                    summary.Children.Add(new TextBlock { Text = "Renamed files:", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
                    foreach (var s in successes)
                    {
                        summary.Children.Add(new TextBlock { Text = $"{s.Original} -> {s.Destination}" });
                    }

                    var ok = new ContentDialog
                    {
                        Title = "Rename complete",
                        Content = summary,
                        CloseButtonText = "OK",
                        XamlRoot = Content.XamlRoot
                    };

                    try
                    {
                        if (MyCheckBox != null)
                        {
                            MyCheckBox.IsChecked = true;
                        }
                    }
                    catch
                    {
                        // ignore
                    }

                    _ = ok.ShowAsync();
                }
                else
                {
                    var summaryText = new System.Text.StringBuilder();
                    summaryText.AppendLine("Some files could not be renamed:");
                    foreach (var f in failures)
                    {
                        summaryText.AppendLine($"{f.Original} -> {f.Destination}: {f.Error}");
                    }

                    var err = new ContentDialog
                    {
                        Title = "Rename errors",
                        Content = summaryText.ToString(),
                        CloseButtonText = "OK",
                        XamlRoot = Content.XamlRoot
                    };
                    _ = err.ShowAsync();
                }
            }
        }

        private static string LoadSystemOutput()
        {
            var exeDir = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            var locationFile = Path.Combine(exeDir, "system.location");

            if (!File.Exists(locationFile))
            {
                throw new FileNotFoundException($"Missing {locationFile}");
            }

            string relativePath;
            using (var reader = new StreamReader(locationFile))
            {
                for (int i = 1; i < 4; i++) { reader.ReadLine(); }// skip first 3 lines 

                relativePath = reader.ReadLine()?.Trim();
            }

            if (string.IsNullOrEmpty(relativePath))
            {
                throw new InvalidOperationException(
                    $"system.location [output file] is empty or invalid: {locationFile}");
            }

            var programFilesX86 =
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

            var outputlocation = Path.Combine(programFilesX86, relativePath);

            if (!Directory.Exists(outputlocation))
            {
                throw new DirectoryNotFoundException($"Nivo output folder not found: {outputlocation}");
            }

            return outputlocation;

        }


        // Build the final Arguments string by performing placeholder replacements.
        // plateVar: value to substitute for [PlateVar]
        // hBarcode: value to substitute for [Hbarcode] and [HBarcode]
        // plateIndex: integer to replace '*' placeholders

        //platevar counter for plates launched to assign plate var

        private string BuildCliArguments(string plateVar, string hBarcode, string _selectedPlateText, int cplates)
        {
            string result = string.Empty;

            // Use protocol_Plate (Tag) to select the correct plate letter
            if (!string.IsNullOrEmpty(protocol_Plate) && protocol_Plate.Length >= cplates)
            {
                plateVar = protocol_Plate[cplates - 1].ToString();
            }

            var exeDir = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            var locationFile = Path.Combine(exeDir, "system.location");

            if (!File.Exists(locationFile))
            {
                throw new FileNotFoundException($"Missing {locationFile}");
            }

            string relativePath;
            using (var reader = new StreamReader(locationFile))
            {
                for (int i = 1; i < 6; i++) { reader.ReadLine(); }// skip first 5 lines 
                relativePath = reader.ReadLine()?.Trim();
            }

            if (string.IsNullOrEmpty(relativePath))
            {
                throw new InvalidOperationException(
                    $"system.location [MyAssays Protocol] is empty or invalid: {locationFile}");
            }

            // Replace {plateVar} in the template with the actual plateVar
            relativePath = relativePath.Replace("{plateVar}", plateVar);

            var programFilesX86 =
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

            var MyAssaysProc = Path.Combine(programFilesX86, relativePath);

            if (!File.Exists(MyAssaysProc))
            {
                throw new FileNotFoundException($"Nivo output folder not found: {MyAssaysProc}");
            }



            var argsList = new List<string>(ArgumentTemplate.Length);
            foreach (var fragment in ArgumentTemplate)
            {
                // Replace placeholders
                var replaced = fragment
                    .Replace("[ProtocolPath]", MyAssaysProc)
                    .Replace("[PlateVar]", plateVar)
                    .Replace("[Hbarcode]", hBarcode)
                    // Replace a single '*' occurrence with the plateIndex string
                    .Replace("*", cplates.ToString());

                argsList.Add(replaced);
            }

            result = string.Join(" ", argsList);

            return result;
        }

        // Helper that launches the CLI for a specific plate index.
        private async Task LaunchCliForPlateAsync(int cplates)
        {
            var plate = protocol_Plate ?? string.Empty;
            var barcode = MyTextBox?.Text ?? string.Empty;

            //check for number of plats also required
            if (string.IsNullOrWhiteSpace(plate) || string.IsNullOrWhiteSpace(barcode))
            {
                var err = new ContentDialog
                {
                    Title = "Missing input",
                    Content = "Please select a plate and enter a barcode before launching.",
                    CloseButtonText = "OK",
                    XamlRoot = Content.XamlRoot
                };
                _ = err.ShowAsync();
                return;
            }

            if ((bool)(MyCheckBox.IsChecked != true))
            {
                var err = new ContentDialog
                {
                    Title = "Files not processed.",
                    Content = "Please process files",
                    CloseButtonText = "OK",
                    XamlRoot = Content.XamlRoot
                };
                _ = err.ShowAsync();
                return;
            }



            try
            {
                var plateVar = string.Empty;
                var hBarcode = MyTextBox?.Text?.Trim() ?? string.Empty;

                var args = BuildCliArguments(plateVar, hBarcode, _selectedPlateText, cplates);
                Debug.WriteLine($"CLI Arguments: /C \"\"{CliExePath}\" {args}\""); // Output to Visual Studio debug pane
                var psi = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/C \"\"{CliExePath}\" {args}\"", // /C runs and closes, /K keeps window open
                    UseShellExecute = true,
                    CreateNoWindow = false,
                };

                using var proc = System.Diagnostics.Process.Start(psi);
                if (proc == null)
                {
                    var dlg = new ContentDialog
                    {
                        Title = "Launch failed",
                        Content = "Could not start the process.",
                        CloseButtonText = "OK",
                        XamlRoot = Content.XamlRoot
                    };
                    _ = dlg.ShowAsync();
                    return;
                }

                // Wait for exit off the UI thread
                await Task.Run(() =>
                {
                    try
                    {
                        proc.WaitForExit();
                    }
                    catch
                    {
                        // ignore wait errors
                    }
                });

                var ok = new ContentDialog
                {
                    Title = "Launched",
                    Content = $"My assays launched for plate {cplates}.",
                    CloseButtonText = "OK",
                    XamlRoot = Content.XamlRoot
                };
                _ = ok.ShowAsync();
            }
            catch (Exception ex)
            {
                var dlg = new ContentDialog
                {
                    Title = "Error launching CLI",
                    Content = $"Exception: {ex.Message}",
                    CloseButtonText = "OK",
                    XamlRoot = Content.XamlRoot
                };
                _ = dlg.ShowAsync();
            }
        }


        private async void OnCombineClick(object sender, RoutedEventArgs e)
        {
            var exeDir = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            var locationFile = Path.Combine(exeDir, "system.location");


            var plate = protocol_Plate ?? string.Empty;
            var barcode = MyTextBox?.Text ?? string.Empty;

            //check for number of plats also required
            if (string.IsNullOrWhiteSpace(barcode))
            {
                var err = new ContentDialog
                {
                    Title = "Missing input",
                    Content = "Please enter a barcode before launching.",
                    CloseButtonText = "OK",
                    XamlRoot = Content.XamlRoot
                };
                _ = err.ShowAsync();
                return;
            }

            if (!File.Exists(locationFile))
            {
                var oops = new ContentDialog
                {
                    Title = "Launched",
                    Content = "location file not found.",
                    CloseButtonText = "OK",
                    XamlRoot = Content.XamlRoot
                };
                _ = oops.ShowAsync();
                return;
            }

            string dataOutput;
            using (var reader = new StreamReader(locationFile))
            {
                for (int i = 1; i < 8; i++) { reader.ReadLine(); }// skip first 7 lines 
                dataOutput = reader.ReadLine()?.Trim();
            }

            var hBarcode = MyTextBox?.Text?.Trim() ?? string.Empty;
            var files = Directory.GetFiles(dataOutput, $"{hBarcode}*.csv");

            if (files.Length == 0)
            {
                var oops = new ContentDialog
                {
                    Title = "Launched",
                    Content = "No files matching entered HBarcode",
                    CloseButtonText = "OK",
                    XamlRoot = Content.XamlRoot
                };
                _ = oops.ShowAsync();
                return;
            }



            if (files.Length >= 1)
            {
                var finalDialog = new ContentDialog
                {
                    Title = "Launched",
                    Content = "files to combine: " + string.Join(", ", files),
                    CloseButtonText = "Not my plates",
                    PrimaryButtonText = "Merge and open",
                    XamlRoot = Content.XamlRoot
                };
                var finalResult = await finalDialog.ShowAsync();
                if (finalResult != ContentDialogResult.Primary)
                {
                    return;
                }
            }
            var processedDir = Path.Combine(dataOutput, "Processed");
            if (!Directory.Exists(processedDir))
            {
                Directory.CreateDirectory(processedDir);
            }
            var fileProcessed = Path.Combine(processedDir, $"{hBarcode}_processed.csv");

            var allCsv = Directory.EnumerateFiles(dataOutput, hBarcode + "*.csv", SearchOption.TopDirectoryOnly);
            string[] header = { File.ReadLines(allCsv.First()).First(l => !string.IsNullOrWhiteSpace(l)) };
            var mergedData = allCsv
                .SelectMany(csv => File.ReadLines(csv)
                    .SkipWhile(l => string.IsNullOrWhiteSpace(l)).Skip(1));// skip header of each file     
            File.WriteAllLines(fileProcessed, header.Concat(mergedData));

            Console.WriteLine($"Combined file created: {fileProcessed}");

            FileInfo fi = new FileInfo(fileProcessed);
            if (fi.Exists)
            {
                try
                {

                    string Excel;
                    using (var reader = new StreamReader(locationFile))
                    {
                        for (int i = 1; i < 10; i++) { reader.ReadLine(); }
                        Excel = reader.ReadLine()?.Trim();
                    }

                    ProcessStartInfo startInfo = new ProcessStartInfo();
                    startInfo.FileName = Excel;
                    startInfo.Arguments = fileProcessed;
                    startInfo.UseShellExecute = true;
                    startInfo.WindowStyle = ProcessWindowStyle.Normal;

                    using (Process proc = Process.Start(startInfo)) ;
                }
                catch (Exception)
                {
                    var err = new ContentDialog
                    {
                        Title = "Open Excel",
                        Content = "An error has occoured openeing excel",
                        CloseButtonText = "OK",
                        XamlRoot = Content.XamlRoot
                    };
                    _ = err.ShowAsync();
                    return;
                }
            }
        }

        private async void LaunchCliButton_Click(object sender, RoutedEventArgs e)
        {
 
            if (numberOfPlates <= 0)
            {
                var dlg = new ContentDialog
                {
                    Title = "No plates selected",
                    Content = "Please select the number of plates before launching.",
                    CloseButtonText = "OK",
                    XamlRoot = Content.XamlRoot
                };
                _ = dlg.ShowAsync();
                return;
            }

            // Launch the CLI for each plate index from 1 to numberOfPlates
            for (int cplates = 1; cplates <= numberOfPlates; cplates++)
            {
                await LaunchCliForPlateAsync(cplates);
            }


        }

    }
}