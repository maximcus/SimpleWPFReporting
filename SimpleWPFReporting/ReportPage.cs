using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace SimpleWPFReporting
{
    internal class ReportPage
    {
        private readonly Grid layoutRootGrid;
        private readonly StackPanel stackPanel = new StackPanel();
        private readonly FrameworkElement reportHeader;
        private readonly FrameworkElement reportFooter;

        internal ReportPage(Size reportSize, StackPanel reportContainer, Thickness margin, object dataContext, DataTemplate reportHeaderDataTemplate, DataTemplate reportFooterDataTemplate, int pageNumber)
        {
            layoutRootGrid = new Grid
            {
                Margin = margin,
                Background = reportContainer.Background,
                Resources = reportContainer.Resources,
                DataContext = dataContext,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };

            layoutRootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            layoutRootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            layoutRootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            if (reportHeaderDataTemplate != null)
            {
                reportHeader = reportHeaderDataTemplate.LoadContent() as FrameworkElement;

                if (reportHeader == null)
                    throw new Exception($"Couldn't cast content of {nameof(reportHeaderDataTemplate)} to {nameof(FrameworkElement)}");

                AddPageNumberResource(reportHeader, pageNumber);
                Grid.SetRow(reportHeader, 0);
                layoutRootGrid.Children.Add(reportHeader);
            }

            Grid.SetRow(stackPanel, 1);
            layoutRootGrid.Children.Add(stackPanel);

            if (reportFooterDataTemplate != null)
            {
                reportFooter = reportFooterDataTemplate.LoadContent() as FrameworkElement;

                if (reportFooter == null)
                    throw new Exception($"Couldn't cast content of {nameof(reportFooterDataTemplate)} to {nameof(FrameworkElement)}");

                AddPageNumberResource(reportFooter, pageNumber);
                Grid.SetRow(reportFooter, 2);
                layoutRootGrid.Children.Add(reportFooter);
            }

            layoutRootGrid.Measure(reportSize);
            layoutRootGrid.Arrange(new Rect(reportSize));
            layoutRootGrid.UpdateLayout();
        }

        internal void AddElement(UIElement element) => stackPanel.Children.Add(element);

        internal UIElementCollection Children => layoutRootGrid.Children;

        private static void AddPageNumberResource(FrameworkElement element, int pageNumber)
        {
            if (element.Resources.Contains("PageNumber"))
                element.Resources["PageNumber"] = pageNumber.ToString();
            else
                element.Resources.Add("PageNumber", pageNumber.ToString());
        }

        internal Grid LayoutRoot => layoutRootGrid;

        internal void ClearChildren()
        {
            stackPanel.Children.Clear();
            layoutRootGrid.Children.Clear();
        }

        internal double GetChildrenActualHeight()
        {
            double childrenActualHeight = 
                stackPanel.Children
                          .Cast<FrameworkElement>()
                          .Sum(elm => GetActualHeightPlusMargin(elm));

            if (reportHeader != null)
                childrenActualHeight += GetActualHeightPlusMargin(reportHeader);

            if (reportFooter != null)
                childrenActualHeight += GetActualHeightPlusMargin(reportFooter);

            return childrenActualHeight;
        }

        private static double GetActualHeightPlusMargin(FrameworkElement elm)
        {
            return elm.ActualHeight + elm.Margin.Top + elm.Margin.Bottom;
        }
    }
}
