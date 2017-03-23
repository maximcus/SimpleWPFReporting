# SimpleWPFReporting

Reporting in WPF (XAML) made easy

## How to get it

SimpleWPFReporting package is available on [NuGet](https://www.nuget.org/packages/SimpleWPFReporting/)

You can use Package Manager Console to get it:

~~~powershell
PM> Install-Package SimpleWPFReporting
~~~

## How does it work

This project gives you the ability to create any report with full power of WPF XAML. SimpleWPFReporting will take care of exporting it as PDF or printing it. The important point is that SimpleWPFReporting will not rasterize your report before exporting or printing.

To be able to divide any report into pages this project expects that you will use vertical `StackPanel` as a report body. You should put your report elements into this `StackPanel`. SimpleWPFReporting accepts `StackPanel` as a parameter and automatically divides report elements into pages.

## How to use it

The main API of SimpleWPFReporting consists of `Report` static class, `ExportReportAsPdf` method, and `PrintReport` method. Both of the methods accept the same arguments.

### Required arguments

`StackPanel reportContainer` is your report body containing report elements.

`object dataContext` is the data context of your report.

`Thickness margin` is the margin applied to every report page. Both of the methods have overloads without the margin argument. They use `Thickness(25)` as default margin.

`ReportOrientation orientation` is the orientation (`Portrait` or `Landscape`) of your report. 

### Optional arguments

`ResourceDictionary resourceDictionary` contains all resources used in the report.

`Brush backgroundBrush` is the background brush for every page of the report.

`DataTemplate reportHeaderDataTemplate` is the `DataTemplate` used to create a header for every page of the report.

`bool headerOnlyOnTheFirstPage` allows you to use header only on the first page of the report. (default is false)

`DataTemplate reportFooterDataTemplate`  is the `DataTemplate` used to create a footer for every page of the report.

`bool footerStartsFromTheSecondPage` allows you to not use footer on the first page of the report. (default is false)

### Page number

Every header and footer will be supplied with `PageNumber` DynamicResource. You can use it as you wish. For example, this is simplest possible footer DataTemplate:

```XML
<DataTemplate x:Key="ReportFooterDataTemplate">
	<TextBlock Text="{DynamicResource PageNumber}" HorizontalAlignment="Right"/>
</DataTemplate>
```


