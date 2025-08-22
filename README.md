# How to clear filter while importing excel into WPF Spreadsheet

This repository contains the sample which shows the how to clear the filter while importing sheet into [WPF Spreadsheet](https://www.syncfusion.com/wpf-controls/spreadsheet) (SfSpreadsheet).

`Spreadsheet` control loads the Excel workbook with filter’s if the filter applied before by user. You can clear the filter and load Excel workbook without filter by clearing the filter by setting [worksheet.AutoFilters.FilterRange](https://help.syncfusion.com/cr/wpf/Syncfusion.XlsIO.IAutoFilters.html#Syncfusion_XlsIO_IAutoFilters_FilterRange) property to null in [WorkbookLoaded](https://help.syncfusion.com/cr/wpf/Syncfusion.UI.Xaml.Spreadsheet.SfSpreadsheet.html#Syncfusion_UI_Xaml_Spreadsheet_SfSpreadsheet_WorkbookLoaded) event.

``` csharp
//To clear the filter
spreadsheet.WorkbookLoaded += Spreadsheet_WorkbookLoaded;
 
private void Spreadsheet_WorkbookLoaded(object sender, Syncfusion.UI.Xaml.Spreadsheet.Helpers.WorkbookLoadedEventArgs args)
{
    foreach (var sheet in args.GridCollection)
    {
        if (sheet.Worksheet.AutoFilters.FilterRange != null)
        {
            sheet.Worksheet.AutoFilters.FilterRange = null;
        }
    }
}
```

![](https://www.syncfusion.com/uploads/user/kb/wpf/wpf-48664/wpf-48664_img1.png)