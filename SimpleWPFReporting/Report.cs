using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Printing;
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

#pragma warning disable 1591
namespace SimpleWPFReporting
{
    public static class Report
    {
        const int DIUPerInch = 96;
        private static readonly Thickness defaultMargin = new Thickness(25);

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
        /// Divides elements of reportContainer into pages and prints them
        /// </summary>
        /// <param name="reportContainer">StackPanel containing report elements</param>
        /// <param name="dataContext">Data Context used in the report</param>
        /// <param name="margin">Margin of a report page</param>
        /// <param name="orientation">Landscape or Portrait orientation</param>
        /// <param name="resourceDictionary">Resources used in report</param>
        /// <param name="backgroundBrush">Brush that will be used as background for report page</param>
        /// <param name="reportHeaderDataTemplate">
        /// Optional header for each page
        /// Note: You can use DynamicResource PageNumber in this template to display page number
        /// </param>
        /// <param name="headerOnlyOnTheFirstPage">Use header only on the first page (default is false)</param>
        /// <param name="reportFooterDataTemplate">
        /// Optional footer for each page
        /// Note: You can use DynamicResource PageNumber in this template to display page number
        /// </param>
        /// <param name="footerStartsFromTheSecondPage">Do not use footer on the first page (default is false)</param> 
        public static void PrintReport(
            StackPanel reportContainer, 
            object dataContext, 
            Thickness margin, 
            ReportOrientation orientation,
            ResourceDictionary resourceDictionary = null,
            Brush backgroundBrush = null,
            DataTemplate reportHeaderDataTemplate = null,
            bool headerOnlyOnTheFirstPage = false,
            DataTemplate reportFooterDataTemplate = null,
            bool footerStartsFromTheSecondPage = false)
        {
            PrintDialog printDialog = new PrintDialog();

            bool? result = printDialog.ShowDialog();

            if (result != true) return;

            Size reportSize = GetReportSize(reportContainer, margin, orientation, printDialog);

            List<FrameworkElement> ReportElements = new List<FrameworkElement>(reportContainer.Children.Cast<FrameworkElement>());
            reportContainer.Children.Clear(); //to avoid exception "Specified element is already the logical child of another element."

            List<ReportPage> ReportPages = 
                GetReportPages(
                    resourceDictionary,
                    backgroundBrush,
                    ReportElements, 
                    dataContext, 
                    margin, 
                    reportSize, 
                    reportHeaderDataTemplate, 
                    headerOnlyOnTheFirstPage, 
                    reportFooterDataTemplate, 
                    footerStartsFromTheSecondPage);

            try
            {
                ReportPages.ForEach(reportPage => reportPage.Scale(reportSize, printDialog));
                ReportPages.ForEach((reportPage, index) => printDialog.PrintVisual(reportPage.LayoutRoot, $"Карточка Точки {index + 1}"));
            }
            finally
            {
                ReportPages.ForEach(reportPage => reportPage.ClearChildren());
                ReportElements.ForEach(elm => reportContainer.Children.Add(elm));
                reportContainer.UpdateLayout();
            }
        }

        /// <summary>
        /// Divides elements of reportContainer into pages and prints them
        /// NOTE: this overload uses default margin of 25
        /// </summary>
        /// <param name="reportContainer">StackPanel containing report elements</param>
        /// <param name="dataContext">Data Context used in the report</param>
        /// <param name="orientation">Landscape or Portrait orientation</param>
        /// <param name="resourceDictionary">Resources used in report</param>
        /// <param name="backgroundBrush">Brush that will be used as background for report page</param>
        /// <param name="reportHeaderDataTemplate">
        /// Optional header for each page
        /// Note: You can use DynamicResource PageNumber in this template to display page number
        /// </param>
        /// <param name="headerOnlyOnTheFirstPage">Use header only on the first page (default is false)</param>
        /// <param name="reportFooterDataTemplate">
        /// Optional footer for each page
        /// Note: You can use DynamicResource PageNumber in this template to display page number
        /// </param>
        /// <param name="footerStartsFromTheSecondPage">Do not use footer on the first page (default is false)</param> 
        public static void PrintReport(
            StackPanel reportContainer,
            object dataContext,
            ReportOrientation orientation,
            ResourceDictionary resourceDictionary = null,
            Brush backgroundBrush = null,
            DataTemplate reportHeaderDataTemplate = null,
            bool headerOnlyOnTheFirstPage = false,
            DataTemplate reportFooterDataTemplate = null,
            bool footerStartsFromTheSecondPage = false)
        {
            PrintReport(
                reportContainer, 
                dataContext, 
                defaultMargin, 
                orientation,
                resourceDictionary,
                backgroundBrush,
                reportHeaderDataTemplate, 
                headerOnlyOnTheFirstPage, 
                reportFooterDataTemplate, 
                footerStartsFromTheSecondPage);
        }

        /// <summary>
        /// Divides elements of reportContainer into pages and exports them as PDF
        /// </summary>
        /// <param name="reportContainer">StackPanel containing report elements</param>
        /// <param name="dataContext">Data Context used in the report</param> 
        /// <param name="margin">Margin of a report page</param>
        /// <param name="orientation">Landscape or Portrait orientation</param>
        /// <param name="resourceDictionary">Resources used in report</param>
        /// <param name="backgroundBrush">Brush that will be used as background for report page</param>
        /// <param name="reportHeaderDataTemplate">
        /// Optional header for each page
        /// Note: You can use DynamicResource PageNumber in this template to display page number
        /// </param>
        /// <param name="headerOnlyOnTheFirstPage">Use header only on the first page (default is false)</param>
        /// <param name="reportFooterDataTemplate">
        /// Optional footer for each page
        /// Note: You can use DynamicResource PageNumber in this template to display page number
        /// </param>
        /// <param name="footerStartsFromTheSecondPage">Do not use footer on the first page (default is false)</param> 
        public static void ExportReportAsPdf(
            StackPanel reportContainer, 
            object dataContext, 
            Thickness margin, 
            ReportOrientation orientation,
            ResourceDictionary resourceDictionary = null,
            Brush backgroundBrush = null,
            DataTemplate reportHeaderDataTemplate = null,
            bool headerOnlyOnTheFirstPage = false,
            DataTemplate reportFooterDataTemplate = null,
            bool footerStartsFromTheSecondPage = false)
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

            List<ReportPage> ReportPages = 
                GetReportPages(
                    resourceDictionary,
                    backgroundBrush,
                    ReportElements, 
                    dataContext, 
                    margin, 
                    reportSize,
                    reportHeaderDataTemplate, 
                    headerOnlyOnTheFirstPage, 
                    reportFooterDataTemplate, 
                    footerStartsFromTheSecondPage);

            FixedDocument fixedDocument = new FixedDocument();

            try
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    System.IO.Packaging.Package package = System.IO.Packaging.Package.Open(memoryStream, FileMode.Create);
                    XpsDocument xpsDocument = new XpsDocument(package);
                    XpsDocumentWriter xpsDocumentWriter = XpsDocument.CreateXpsDocumentWriter(xpsDocument);

                    foreach (Grid reportPage in ReportPages.Select(reportPage => reportPage.LayoutRoot))
                    {
                        reportPage.Width = reportPage.ActualWidth;
                        reportPage.Height = reportPage.ActualHeight;

                        FixedPage newFixedPage = new FixedPage();
                        newFixedPage.Children.Add(reportPage);
                        newFixedPage.Measure(reportSize);
                        newFixedPage.Arrange(new Rect(reportSize));
                        newFixedPage.Width = newFixedPage.ActualWidth;
                        newFixedPage.Height = newFixedPage.ActualHeight;
                        newFixedPage.Background = backgroundBrush;
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
                ReportPages.ForEach(reportPage => reportPage.ClearChildren());
                ReportElements.ForEach(elm => reportContainer.Children.Add(elm));
                reportContainer.UpdateLayout();
            }
        }

        /// <summary>
        /// Divides elements of reportContainer into pages and exports them as PDF
        /// NOTE: this overload uses default margin of 25
        /// </summary>
        /// <param name="reportContainer">StackPanel containing report elements</param>
        /// <param name="dataContext">Data Context used in the report</param> 
        /// <param name="orientation">Landscape or Portrait orientation</param>
        /// <param name="resourceDictionary">Resources used in report</param>
        /// <param name="backgroundBrush">Brush that will be used as background for report page</param>
        /// <param name="reportHeaderDataTemplate">
        /// Optional header for each page
        /// Note: You can use DynamicResource PageNumber in this template to display page number
        /// </param>
        /// <param name="headerOnlyOnTheFirstPage">Use header only on the first page (default is false)</param>
        /// <param name="reportFooterDataTemplate">
        /// Optional footer for each page
        /// Note: You can use DynamicResource PageNumber in this template to display page number
        /// </param>
        /// <param name="footerStartsFromTheSecondPage">Do not use footer on the first page (default is false)</param> 
        public static void ExportReportAsPdf(
            StackPanel reportContainer,
            object dataContext,
            ReportOrientation orientation,
            ResourceDictionary resourceDictionary = null,
            Brush backgroundBrush = null,
            DataTemplate reportHeaderDataTemplate = null,
            bool headerOnlyOnTheFirstPage = false,
            DataTemplate reportFooterDataTemplate = null,
            bool footerStartsFromTheSecondPage = false)
        {
            ExportReportAsPdf(
                reportContainer, 
                dataContext, 
                defaultMargin, 
                orientation,
                resourceDictionary,
                backgroundBrush,
                reportHeaderDataTemplate, 
                headerOnlyOnTheFirstPage,
                reportFooterDataTemplate, 
                footerStartsFromTheSecondPage);
        }

        private static List<ReportPage> GetReportPages(
            ResourceDictionary resourceDictionary,
            Brush backgroundBrush,
            List<FrameworkElement> ReportElements, 
            object dataContext, 
            Thickness margin, 
            Size reportSize,
            DataTemplate reportHeaderDataTemplate,
            bool headerOnlyOnTheFirstPage,
            DataTemplate reportFooterDataTemplate,
            bool footerStartsFromTheSecondPage)
        {
            int pageNumber = 1;

            List<ReportPage> ReportPages = 
                new List<ReportPage>
                {
                    new ReportPage(
                        reportSize,
                        backgroundBrush, 
                        margin, 
                        dataContext,
                        resourceDictionary,
                        reportHeaderDataTemplate,
                        (footerStartsFromTheSecondPage) ? null : reportFooterDataTemplate, 
                        pageNumber)
                };

            foreach (FrameworkElement reportVisualElement in ReportElements)
            {
                if (ReportPages.Last().GetChildrenActualHeight() + GetActualHeightPlusMargin(reportVisualElement) > reportSize.Height - margin.Top - margin.Bottom)
                {
                    pageNumber++;

                    ReportPages.Add(
                        new ReportPage(
                            reportSize,
                            backgroundBrush, 
                            margin, 
                            dataContext,
                            resourceDictionary,
                            (headerOnlyOnTheFirstPage) ? null : reportHeaderDataTemplate, 
                            reportFooterDataTemplate, 
                            pageNumber));
                }

                ReportPages.Last().AddElement(reportVisualElement);
            }

            foreach (ReportPage reportPage in ReportPages)
            {
                reportPage.LayoutRoot.Measure(reportSize);
                reportPage.LayoutRoot.Arrange(new Rect(reportSize));
                reportPage.LayoutRoot.UpdateLayout();
            }

            return ReportPages;
        }

        private static double GetActualHeightPlusMargin(FrameworkElement elm)
        {
            return elm.ActualHeight + elm.Margin.Top + elm.Margin.Bottom;
        }

        private static Size GetReportSize(StackPanel reportContainer, Thickness margin, ReportOrientation orientation, PrintDialog printDialog = null)
        {
            if (printDialog == null)
                printDialog = new PrintDialog();

            double reportWidth = reportContainer.ActualWidth + margin.Left + margin.Right;

            if (orientation == ReportOrientation.Landscape)
                printDialog.PrintTicket.PageOrientation = PageOrientation.Landscape;

            double reportHeight = (reportWidth / printDialog.PrintableAreaWidth) * printDialog.PrintableAreaHeight;

            return new Size(reportWidth, reportHeight);
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
#pragma warning restore 1591


