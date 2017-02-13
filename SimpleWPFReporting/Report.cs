using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Xps;
using System.Windows.Xps.Packaging;
using Microsoft.Win32;
using PdfSharp.Xps;

namespace SimpleWPFReporting
{
    public static class Report
    {
        const int DIUPerInch = 96;
        public static void ExportVisualAsXps(Visual visual)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                DefaultExt = ".xps",
                Filter = "XPS Documents (.xps)|*.xps"
            };

            bool? result = saveFileDialog.ShowDialog();

            if (result != true) return;

            XpsDocument xpsDocument = new XpsDocument(saveFileDialog.FileName, FileAccess.Write);
            XpsDocumentWriter xpsDocumentWriter = XpsDocument.CreateXpsDocumentWriter(xpsDocument);

            xpsDocumentWriter.Write(visual);
            xpsDocument.Close();
        }

        public static void ExportVisualAsPdf(Visual visual)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                DefaultExt = ".pdf",
                Filter = "PDF Documents (.pdf)|*.pdf"
            };

            bool? result = saveFileDialog.ShowDialog();

            if (result != true) return;

            using (MemoryStream memoryStream = new MemoryStream())
            {
                System.IO.Packaging.Package package = System.IO.Packaging.Package.Open(memoryStream, FileMode.Create);
                XpsDocument xpsDocument = new XpsDocument(package);
                XpsDocumentWriter xpsDocumentWriter = XpsDocument.CreateXpsDocumentWriter(xpsDocument);

                xpsDocumentWriter.Write(visual);
                xpsDocument.Close();
                package.Close();

                var pdfXpsDoc = PdfSharp.Xps.XpsModel.XpsDocument.Open(memoryStream);
                XpsConverter.Convert(pdfXpsDoc, saveFileDialog.FileName, 0);
            }
        }

        private static readonly Lazy<int> dpiX = new Lazy<int>(() =>
        {
            PropertyInfo dpiXProperty = typeof(SystemParameters).GetProperty("DpiX", BindingFlags.NonPublic | BindingFlags.Static);
            return (int)dpiXProperty.GetValue(null, null);
        });

        private static readonly Lazy<int> dpiY = new Lazy<int>(() =>
        {
            var dpiYProperty = typeof(SystemParameters).GetProperty("Dpi", BindingFlags.NonPublic | BindingFlags.Static);
            return (int)dpiYProperty.GetValue(null, null);
        });

        private static void SaveFrameworkElementAsImage(FrameworkElement element, string filePath, BitmapEncoder bitmapEncoder)
        {
            RenderTargetBitmap bitmap = new RenderTargetBitmap(
                pixelWidth: Convert.ToInt32((element.ActualWidth / DIUPerInch) * dpiX.Value), 
                pixelHeight: Convert.ToInt32((element.ActualHeight / DIUPerInch) * dpiY.Value), 
                dpiX: dpiX.Value, 
                dpiY: dpiY.Value, 
                pixelFormat: PixelFormats.Pbgra32);
            bitmap.Render(element);

            bitmapEncoder.Frames.Add(BitmapFrame.Create(bitmap));
            using (Stream fs = File.Create(filePath))
            {
                bitmapEncoder.Save(fs);
            }
        }

        public static void ExportFrameworkElementAsJpg(FrameworkElement element)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                DefaultExt = ".jpg",
                Filter = "JPG Images (.jpg)|*.jpg"
            };

            bool? result = saveFileDialog.ShowDialog();

            if (result != true) return;

            SaveFrameworkElementAsImage(element, saveFileDialog.FileName, new JpegBitmapEncoder {QualityLevel = 100});
        }

        public static void ExportFrameworkElementAsBmp(FrameworkElement element)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                DefaultExt = ".bmp",
                Filter = "BMP Images (.bmp)|*.bmp"
            };

            bool? result = saveFileDialog.ShowDialog();

            if (result != true) return;

            SaveFrameworkElementAsImage(element, saveFileDialog.FileName, new BmpBitmapEncoder());
        }

        public static void ExportFrameworkElementAsPng(FrameworkElement element)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                DefaultExt = ".png",
                Filter = "PNG Images (.png)|*.png"
            };

            bool? result = saveFileDialog.ShowDialog();

            if (result != true) return;

            SaveFrameworkElementAsImage(element, saveFileDialog.FileName, new PngBitmapEncoder());
        }

        /// <summary>
        /// Prints report using reportContainer
        /// </summary>
        /// <param name="reportContainer">StackPanel containing report elements to be printed sequentially</param>
        /// <param name="dataContext">Data Context used in report</param> 
        /// <param name="margin">By default you should use 
        ///     reportContainer.Margin.Left = reportContainer.Margin.Right to preserve
        ///     report initial dimensions
        /// </param>
        /// <param name="orientation">Landscape or Portrait orientation</param> 
        /// <param name="reportHeaderDataTemplate">Optional header for each page</param>
        public static void PrintReport(StackPanel reportContainer, object dataContext, double margin, ReportOrientation orientation, DataTemplate reportHeaderDataTemplate = null)
        {
            PrintDialog printDialog = new PrintDialog();

            bool? result = printDialog.ShowDialog();

            if (result != true) return;

            Size reportSize = GetReportSize(reportContainer, margin, orientation, printDialog);

            List<FrameworkElement> ReportElements = new List<FrameworkElement>(reportContainer.Children.Cast<FrameworkElement>());
            reportContainer.Children.Clear(); //to avoid exception "Specified element is already the logical child of another element."

            List<StackPanel> ReportPages = GetReportPages(reportHeaderDataTemplate, reportContainer, margin, ReportElements, reportSize, dataContext);

            try
            {
                ReportPages.ForEach((reportPage, index) => printDialog.PrintVisual(reportPage, $"Карточка Точки {index + 1}"));
            }
            finally
            {
                ReportPages.ForEach(reportPage => reportPage.Children.Clear());
                ReportElements.ForEach(elm => reportContainer.Children.Add(elm));
                reportContainer.UpdateLayout();
            }
        }

        /// <summary>
        /// Export as report to PDF
        /// </summary>
        /// <param name="reportContainer">StackPanel containing report elements to be split into pages sequentially</param>
        /// <param name="dataContext">Data Context used in report</param> 
        /// <param name="margin">By default you should use 
        ///     reportContainer.Margin.Left = reportContainer.Margin.Right to preserve
        ///     report initial dimensions
        /// </param>
        /// <param name="orientation">Landscape or Portrait orientation</param>
        /// <param name="reportHeaderDataTemplate">Optional header for each page</param>
        public static void ExportReportAsPdf(StackPanel reportContainer, object dataContext, double margin, ReportOrientation orientation, DataTemplate reportHeaderDataTemplate = null)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                DefaultExt = ".pdf",
                Filter = "PDF Documents (.pdf)|*.pdf"
            };

            bool? result = saveFileDialog.ShowDialog();

            if (result != true) return;

            Size reportSize = GetReportSize(reportContainer, margin, orientation);

            List<FrameworkElement> ReportElements = new List<FrameworkElement>(reportContainer.Children.Cast<FrameworkElement>());
            reportContainer.Children.Clear(); //to avoid exception "Specified element is already the logical child of another element."

            List<StackPanel> ReportPages = GetReportPages(reportHeaderDataTemplate, reportContainer, margin, ReportElements, reportSize, dataContext);

            FixedDocument fixedDocument = new FixedDocument();

            try
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    System.IO.Packaging.Package package = System.IO.Packaging.Package.Open(memoryStream, FileMode.Create);
                    XpsDocument xpsDocument = new XpsDocument(package);
                    XpsDocumentWriter xpsDocumentWriter = XpsDocument.CreateXpsDocumentWriter(xpsDocument);

                    foreach (StackPanel reportPage in ReportPages)
                    {
                        reportPage.Width = reportPage.ActualWidth;

                        FixedPage newFixedPage = new FixedPage();
                        newFixedPage.Children.Add(reportPage);
                        newFixedPage.Measure(reportSize);
                        newFixedPage.Arrange(new Rect(reportSize));
                        newFixedPage.Width = newFixedPage.ActualWidth;
                        newFixedPage.Height = newFixedPage.ActualHeight;
                        newFixedPage.UpdateLayout();

                        PageContent pageContent = new PageContent();
                        ((IAddChild)pageContent).AddChild(newFixedPage);

                        fixedDocument.Pages.Add(pageContent);
                    }

                    xpsDocumentWriter.Write(fixedDocument);
                    xpsDocument.Close();
                    package.Close();

                    var pdfXpsDoc = PdfSharp.Xps.XpsModel.XpsDocument.Open(memoryStream);
                    XpsConverter.Convert(pdfXpsDoc, saveFileDialog.FileName, 0);
                }
            }
            finally
            {
                ReportPages.ForEach(reportPage => reportPage.Children.Clear());
                ReportElements.ForEach(elm => reportContainer.Children.Add(elm));
                reportContainer.UpdateLayout();
            }
        }

        private static List<StackPanel> GetReportPages(DataTemplate reportHeaderDataTemplate, StackPanel reportContainer, double margin, List<FrameworkElement> ReportElements, Size reportSize, object dataContext)
        {
            List<StackPanel> ReportPages = new List<StackPanel> {GetReportPageContainer(reportHeaderDataTemplate, reportContainer, margin, dataContext) };

            foreach (FrameworkElement reportVisualElement in ReportElements)
            {
                if (ReportPages.Last().Children
                    .Cast<FrameworkElement>()
                    .Sum(elm => elm.ActualHeight) + reportVisualElement.ActualHeight > reportSize.Height - margin*3)
                {
                    ReportPages.Add(GetReportPageContainer(reportHeaderDataTemplate, reportContainer, margin, dataContext));
                }

                ReportPages.Last().Children.Add(reportVisualElement);
            }

            foreach (StackPanel reportPage in ReportPages)
            {
                reportPage.Measure(reportSize);
                reportPage.Arrange(new Rect(reportSize));
                reportPage.UpdateLayout();
            }

            return ReportPages;
        }

        private static Size GetReportSize(StackPanel reportContainer, double margin, ReportOrientation orientation, PrintDialog printDialog = null)
        {
            if (printDialog == null)
                printDialog = new PrintDialog();

            double reportWidth = reportContainer.ActualWidth + margin * 2;

            double reportHeight;
            if (orientation == ReportOrientation.Portrait)
                reportHeight = (reportWidth / printDialog.PrintableAreaWidth) * printDialog.PrintableAreaHeight;
            else
                reportHeight = (reportWidth / printDialog.PrintableAreaHeight) * printDialog.PrintableAreaWidth;

            return new Size(reportWidth, reportHeight);
        }

        private static StackPanel GetReportPageContainer(DataTemplate reportHeaderDataTemplate, StackPanel reportContainer, double margin, object dataContext)
        {
            StackPanel reportPageContainer = new StackPanel
            {
                Margin = new Thickness(margin),
                Background = reportContainer.Background,
                Resources = reportContainer.Resources,
                DataContext = dataContext
            };

            if (reportHeaderDataTemplate != null)
            {
                UIElement reportHeader = reportHeaderDataTemplate.LoadContent() as UIElement;

                if (reportHeader != null)
                    reportPageContainer.Children.Add(reportHeader);
                else throw new Exception($"Couldn't cast content of {nameof(reportHeaderDataTemplate)} to {nameof(UIElement)}");
            }

            return reportPageContainer;
        }

        private static void ForEach<T>(this IEnumerable<T> enumeration, Action<T, int> action)
        {
            if (enumeration == null)
                return;

            int index = 0;
            foreach (T item in enumeration)
            {
                action(item, index);
                index++;
            }
        }
    }
}

