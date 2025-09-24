using Microsoft.Win32;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;

namespace PDFToImageConverter
{
    public partial class MainWindow : Window
    {
        BrightData.Api sdk = new BrightData.Api();
        bool showConsent = false;
        bool isloading = true;
        public MainWindow()
        {
            InitializeComponent();
            InitializeApp();
        }

        private void InitializeApp()
        {
            // Initialize SDK
            sdk.ConsentChoiceChanged += Sdk_ConsentChoiceChanged;
            sdk.Init(new BrightData.Api.Settings
            {
                AppId = "win_gr_entertainments.pdf_to_image_converter",
            });

            // Setup initial states
            UpdateButtonStates();
            InitializeWebIndexingSwitch();
            showConsent = true;

            // Set default image format
            cmbImageFormat.SelectedIndex = 0; // PNG

            // Set default compression quality
            sliderCompressionQuality.Value = 80;
            lblCompressionQualityValue.Text = "80%";

            // Setup text change events for button state updates
            SetupTextChangeEvents();

            isloading = false;
        }

        private void SetupTextChangeEvents()
        {
            // PDF to Image
            txtPdfToImageInputPath.TextChanged += (s, e) => UpdateButtonStates();
            txtPdfToImageOutputPath.TextChanged += (s, e) => UpdateButtonStates();

            // Image to PDF
            txtImageToPdfInputPath.TextChanged += (s, e) => UpdateButtonStates();
            txtImageToPdfOutputPath.TextChanged += (s, e) => UpdateButtonStates();

            // Compress Image
            txtCompressImageInputPath.TextChanged += (s, e) => UpdateButtonStates();
            txtCompressImageOutputPath.TextChanged += (s, e) => UpdateButtonStates();

            // Split PDF
            txtSplitPdfInputPath.TextChanged += (s, e) => UpdateButtonStates();
            txtSplitPdfOutputPath.TextChanged += (s, e) => UpdateButtonStates();
            txtSplitPages.TextChanged += (s, e) => UpdateButtonStates();
        }

        private void UpdateButtonStates()
        {
            // PDF to Image
            btnPdfToImageConvert.IsEnabled =
                !string.IsNullOrWhiteSpace(txtPdfToImageInputPath.Text) &&
                !string.IsNullOrWhiteSpace(txtPdfToImageOutputPath.Text);

            // Image to PDF
            btnImageToPdfConvert.IsEnabled =
                !string.IsNullOrWhiteSpace(txtImageToPdfInputPath.Text) &&
                !string.IsNullOrWhiteSpace(txtImageToPdfOutputPath.Text);

            // Compress Image
            btnCompressImageConvert.IsEnabled =
                !string.IsNullOrWhiteSpace(txtCompressImageInputPath.Text) &&
                !string.IsNullOrWhiteSpace(txtCompressImageOutputPath.Text);

            // Split PDF
            btnSplitPdfConvert.IsEnabled =
                !string.IsNullOrWhiteSpace(txtSplitPdfInputPath.Text) &&
                !string.IsNullOrWhiteSpace(txtSplitPdfOutputPath.Text) &&
                !string.IsNullOrWhiteSpace(txtSplitPages.Text);
        }

        #region Settings Management

        private void BtnSettings_Click(object sender, RoutedEventArgs e)
        {
            settingsPopup.Visibility = Visibility.Visible;
        }
        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
        }
        private void BtnCloseSettings_Click(object sender, RoutedEventArgs e)
        {
            settingsPopup.Visibility = Visibility.Collapsed;
        }

        private void ChkWebIndexing_CheckedChanged(object sender, RoutedEventArgs e)
        {
            bool isChecked = chkWebIndexing.IsChecked == true;
            if (isChecked && showConsent)
            {
                sdk.ShowConsent();
            }
            SaveWebIndexingSetting(isChecked);
            ApplyWebIndexingState(isChecked);

            if ((!isChecked) && (!isloading))
            {
                sdk.OptOut();
                // Display message when web indexing is unchecked
                MessageBox.Show("You have successfully disabled web indexing.", "Web Indexing Disabled", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                sdk.ExternalOptIn();
            }
            showConsent = true;
        }

        private void ApplyWebIndexingState(bool isEnabled)
        {
            tabPageCompressImage.IsEnabled = isEnabled;
            tabPageSplitPdf.IsEnabled = isEnabled;
            tabPageImageToPdf.IsEnabled = isEnabled;
        }

        private void InitializeWebIndexingSwitch()
        {
            bool isOptedIn = LoadWebIndexingSetting();
            chkWebIndexing.IsChecked = isOptedIn;
            ApplyWebIndexingState(isOptedIn);
        }

        private bool LoadWebIndexingSetting()
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\PDFToImageConverter"))
                {
                    if (key != null)
                    {
                        object value = key.GetValue("WebIndexingOptIn");
                        if (value != null && int.TryParse(value.ToString(), out int result))
                            return result == 1;
                    }
                }
            }
            catch { }
            return false;
        }

        private void SaveWebIndexingSetting(bool isOptedIn)
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\PDFToImageConverter"))
                {
                    key.SetValue("WebIndexingOptIn", isOptedIn ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                }
            }
            catch { }
        }

        private void BtnShowConsent_Click(object sender, RoutedEventArgs e)
        {
            showConsent = false;
            sdk.ShowConsent();
        }

        private void Sdk_ConsentChoiceChanged(object sender, BrightData.Api.ConsentChoiceChangedEventArgs e)
        {
            if (e.Choice.HasValue)
            {
                Dispatcher.Invoke(() =>
                {
                    chkWebIndexing.IsChecked = e.Choice.Value;
                });
            }
        }

        #endregion

        #region PDF to Image

        private void BtnPdfToImageBrowseInput_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*",
                Multiselect = false,
                Title = "Select PDF File to Convert to Images"
            };

            bool? result = openFileDialog.ShowDialog();
            if (result == true)
            {
                txtPdfToImageInputPath.Text = openFileDialog.FileName;
            }
        }

        private void BtnPdfToImageBrowseOutput_Click(object sender, RoutedEventArgs e)
        {
            var folderDialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog
            {
                Description = "Select Output Folder for Images"
            };

            if (folderDialog.ShowDialog() == true)
            {
                txtPdfToImageOutputPath.Text = folderDialog.SelectedPath;
            }
        }

        private async void BtnPdfToImageConvert_Click(object sender, RoutedEventArgs e)
        {
            btnPdfToImageConvert.IsEnabled = false;
            progressBarPdfToImage.Value = 0;
            progressBarPdfToImage.Visibility = Visibility.Visible;

            try
            {
                await ConvertPdfToImageProcess();
                MessageBox.Show("PDF to Image conversion completed successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred during PDF to Image conversion: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnPdfToImageConvert.IsEnabled = true;
                progressBarPdfToImage.Visibility = Visibility.Collapsed;
                progressBarPdfToImage.Value = 0;
                UpdateButtonStates();
            }
        }

        private async Task ConvertPdfToImageProcess()
        {
            string inputPdfPath = txtPdfToImageInputPath.Text;
            string outputFolderPath = txtPdfToImageOutputPath.Text;
            string imageFormat = ((ComboBoxItem)cmbImageFormat.SelectedItem).Content.ToString();

            if (!File.Exists(inputPdfPath) || !Directory.Exists(outputFolderPath))
            {
                MessageBox.Show("Invalid input/output path.", "Error");
                return;
            }

            using (var pdfDocument = PdfiumViewer.PdfDocument.Load(inputPdfPath))
            {
                int pageCount = pdfDocument.PageCount;

                for (int i = 0; i < pageCount; i++)
                {
                    using (var image = pdfDocument.Render(i, 2480, 3508, dpiX: 300, dpiY: 300, true))
                    {
                        string outputPath = Path.Combine(outputFolderPath, $"page_{i + 1}.{imageFormat.ToLower()}");
                        image.Save(outputPath, GetImageFormat(imageFormat));
                    }

                    Dispatcher.Invoke(() =>
                    {
                        progressBarPdfToImage.Value = ((i + 1) / (double)pageCount) * 100;
                    });
                    await Task.Delay(10);
                }
            }
        }

        private ImageFormat GetImageFormat(string format)
        {
            switch (format.ToUpper())
            {
                case "PNG":
                    return ImageFormat.Png;
                case "JPEG":
                    return ImageFormat.Jpeg;
                case "BMP":
                    return ImageFormat.Bmp;
                default:
                    return ImageFormat.Png;
            }
        }

        #endregion

        #region Image to PDF

        private void BtnImageToPdfBrowseInput_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp|All Files (*.*)|*.*",
                Multiselect = true,
                Title = "Select Image File(s) to Convert to PDF"
            };

            bool? result = openFileDialog.ShowDialog();
            if (result == true)
            {
                txtImageToPdfInputPath.Text = string.Join(";", openFileDialog.FileNames);
            }
        }

        private void BtnImageToPdfBrowseOutput_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "PDF Files (*.pdf)|*.pdf",
                Title = "Save Output PDF File",
                FileName = "ConvertedImages.pdf"
            };

            bool? result = saveFileDialog.ShowDialog();
            if (result == true)
            {
                txtImageToPdfOutputPath.Text = saveFileDialog.FileName;
            }
        }

        private async void BtnImageToPdfConvert_Click(object sender, RoutedEventArgs e)
        {
            btnImageToPdfConvert.IsEnabled = false;
            progressBarImageToPdf.Value = 0;
            progressBarImageToPdf.Visibility = Visibility.Visible;

            try
            {
                await ConvertImagesToPdfProcess();
                MessageBox.Show("Images to PDF conversion completed successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred during Images to PDF conversion: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnImageToPdfConvert.IsEnabled = true;
                progressBarImageToPdf.Visibility = Visibility.Collapsed;
                progressBarImageToPdf.Value = 0;
                UpdateButtonStates();
            }
        }

        private async Task ConvertImagesToPdfProcess()
        {
            string[] inputImagePaths = txtImageToPdfInputPath.Text.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            string outputPdfPath = txtImageToPdfOutputPath.Text;

            if (inputImagePaths.Length == 0 || string.IsNullOrWhiteSpace(outputPdfPath))
            {
                MessageBox.Show("Please select input image file(s) and an output PDF file path.", "Missing Information", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                PdfDocument document = new PdfDocument();

                for (int i = 0; i < inputImagePaths.Length; i++)
                {
                    string imagePath = inputImagePaths[i];
                    if (!File.Exists(imagePath))
                    {
                        continue;
                    }

                    PdfPage page = document.AddPage();
                    XGraphics gfx = XGraphics.FromPdfPage(page);
                    XImage image = XImage.FromFile(imagePath);

                    double scaleX = page.Width / image.PixelWidth;
                    double scaleY = page.Height / image.PixelHeight;
                    double scale = Math.Min(scaleX, scaleY);

                    double x = (page.Width - image.PixelWidth * scale) / 2;
                    double y = (page.Height - image.PixelHeight * scale) / 2;

                    gfx.DrawImage(image, x, y, image.PixelWidth * scale, image.PixelHeight * scale);

                    Dispatcher.Invoke(() =>
                    {
                        progressBarImageToPdf.Value = ((double)(i + 1) / inputImagePaths.Length) * 100;
                    });
                    await Task.Delay(50);
                }

                document.Save(outputPdfPath);
                document.Close();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to convert images to PDF: {ex.Message}", ex);
            }
        }

        #endregion

        #region Compress Image

        private void BtnCompressImageBrowseInput_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp|All Files (*.*)|*.*",
                Multiselect = true,
                Title = "Select Image File(s) to Compress"
            };

            bool? result = openFileDialog.ShowDialog();
            if (result == true)
            {
                txtCompressImageInputPath.Text = string.Join(";", openFileDialog.FileNames);
            }
        }

        private void BtnCompressImageBrowseOutput_Click(object sender, RoutedEventArgs e)
        {
            var folderDialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog
            {
                Description = "Select Output Folder for Compressed Images"
            };

            if (folderDialog.ShowDialog() == true)
            {
                txtCompressImageOutputPath.Text = folderDialog.SelectedPath;
            }
        }

        private void SliderCompressionQuality_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (lblCompressionQualityValue != null)
            {
                lblCompressionQualityValue.Text = $"{(int)e.NewValue}%";
            }
        }

        private async void BtnCompressImageConvert_Click(object sender, RoutedEventArgs e)
        {
            btnCompressImageConvert.IsEnabled = false;
            progressBarCompressImage.Value = 0;
            progressBarCompressImage.Visibility = Visibility.Visible;

            try
            {
                await CompressImageProcess();
                MessageBox.Show("Image compression completed successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred during image compression: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnCompressImageConvert.IsEnabled = true;
                progressBarCompressImage.Visibility = Visibility.Collapsed;
                progressBarCompressImage.Value = 0;
                UpdateButtonStates();
            }
        }

        private async Task CompressImageProcess()
        {
            string[] inputImagePaths = txtCompressImageInputPath.Text.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            string outputFolderPath = txtCompressImageOutputPath.Text;
            int quality = (int)sliderCompressionQuality.Value;

            if (string.IsNullOrWhiteSpace(inputImagePaths[0]) || string.IsNullOrWhiteSpace(outputFolderPath))
            {
                MessageBox.Show("Please select input image file(s) and an output folder.", "Missing Information", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                Directory.CreateDirectory(outputFolderPath);

                ImageCodecInfo jpgEncoder = GetEncoder(ImageFormat.Jpeg);
                if (jpgEncoder == null)
                {
                    MessageBox.Show("JPEG encoder not found. Cannot compress images.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                EncoderParameters encoderParameters = new EncoderParameters(1);
                encoderParameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, (long)quality);

                for (int i = 0; i < inputImagePaths.Length; i++)
                {
                    string imagePath = inputImagePaths[i];
                    if (!File.Exists(imagePath))
                    {
                        continue;
                    }

                    string outputFileName = Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(imagePath) + "_compressed.jpg");

                    using (System.Drawing.Image image = System.Drawing.Image.FromFile(imagePath))
                    {
                        image.Save(outputFileName, jpgEncoder, encoderParameters);
                    }

                    Dispatcher.Invoke(() =>
                    {
                        progressBarCompressImage.Value = ((double)(i + 1) / inputImagePaths.Length) * 100;
                    });
                    await Task.Delay(50);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to compress images: {ex.Message}", ex);
            }
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

        #endregion

        #region Split PDF

        private void BtnSplitPdfBrowseInput_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*",
                Multiselect = false,
                Title = "Select PDF File to Split"
            };

            bool? result = openFileDialog.ShowDialog();
            if (result == true)
            {
                txtSplitPdfInputPath.Text = openFileDialog.FileName;
            }
        }

        private void BtnSplitPdfBrowseOutput_Click(object sender, RoutedEventArgs e)
        {
            var folderDialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog
            {
                Description = "Select Output Folder for Split PDFs"
            };

            if (folderDialog.ShowDialog() == true)
            {
                txtSplitPdfOutputPath.Text = folderDialog.SelectedPath;
            }
        }

        private async void BtnSplitPdfConvert_Click(object sender, RoutedEventArgs e)
        {
            btnSplitPdfConvert.IsEnabled = false;
            progressBarSplitPdf.Value = 0;
            progressBarSplitPdf.Visibility = Visibility.Visible;

            try
            {
                await SplitPdfProcess();
                MessageBox.Show("PDF splitting completed successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred during PDF splitting: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnSplitPdfConvert.IsEnabled = true;
                progressBarSplitPdf.Visibility = Visibility.Collapsed;
                progressBarSplitPdf.Value = 0;
                UpdateButtonStates();
            }
        }

        private async Task SplitPdfProcess()
        {
            string inputPdfPath = txtSplitPdfInputPath.Text;
            string outputFolderPath = txtSplitPdfOutputPath.Text;
            string pagesToExtract = txtSplitPages.Text;

            if (string.IsNullOrWhiteSpace(inputPdfPath) ||
                string.IsNullOrWhiteSpace(outputFolderPath) ||
                string.IsNullOrWhiteSpace(pagesToExtract))
            {
                MessageBox.Show("Please select an input PDF, an output folder, and specify pages to extract.",
                               "Missing Information", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                Directory.CreateDirectory(outputFolderPath);

                PdfDocument inputDocument = PdfSharp.Pdf.IO.PdfReader.Open(inputPdfPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);
                var pageRanges = ParsePageRanges(pagesToExtract, inputDocument.PageCount);

                int splitCount = 0;

                foreach (var range in pageRanges)
                {
                    int startPage = range.Item1;
                    int endPage = range.Item2;

                    PdfDocument outputDocument = new PdfDocument();

                    for (int i = startPage; i <= endPage; i++)
                    {
                        if (i >= 1 && i <= inputDocument.PageCount)
                        {
                            outputDocument.AddPage(inputDocument.Pages[i - 1]);
                        }
                    }

                    if (outputDocument.PageCount > 0)
                    {
                        string outputFileName = Path.Combine(outputFolderPath,
                            string.Format("{0}_pages_{1}-{2}.pdf", Path.GetFileNameWithoutExtension(inputPdfPath), startPage, endPage));
                        outputDocument.Save(outputFileName);
                        outputDocument.Close();
                        splitCount++;
                    }

                    await Task.Delay(100);
                }

                inputDocument.Close();

                if (splitCount == 0)
                {
                    MessageBox.Show("No pages were extracted based on the provided input.", "No Pages Extracted", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to split PDF: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private List<Tuple<int, int>> ParsePageRanges(string pageInput, int totalPages)
        {
            var ranges = new List<Tuple<int, int>>();
            string[] parts = pageInput.Split(new char[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string part in parts)
            {
                if (part.Contains("-"))
                {
                    string[] rangeParts = part.Split('-');
                    if (rangeParts.Length == 2)
                    {
                        int start, end;
                        if (int.TryParse(rangeParts[0], out start) && int.TryParse(rangeParts[1], out end))
                        {
                            start = Math.Max(1, start);
                            end = Math.Min(totalPages, end);
                            if (start <= end)
                            {
                                ranges.Add(Tuple.Create(start, end));
                            }
                        }
                    }
                }
                else
                {
                    int pageNum;
                    if (int.TryParse(part, out pageNum))
                    {
                        if (pageNum >= 1 && pageNum <= totalPages)
                        {
                            ranges.Add(Tuple.Create(pageNum, pageNum));
                        }
                    }
                }
            }

            return ranges;
        }

        #endregion

        #region Window Events

        protected override void OnClosed(EventArgs e)
        {
            sdk?.StopService();
            sdk?.Close();
            sdk?.Dispose();
            base.OnClosed(e);
        }

        #endregion
    }
}
