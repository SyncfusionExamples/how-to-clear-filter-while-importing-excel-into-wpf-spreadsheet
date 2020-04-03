using Syncfusion.Windows.Tools.Controls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Spreadsheet_clearfilter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : RibbonWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            spreadsheet.WorkbookLoaded += Spreadsheet_WorkbookLoaded;
            this.Loaded += MainWindow_Loaded;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                var fileStream = new FileStream(@"..\..\Data\Filtering.xlsx", FileMode.Open);
                spreadsheet.Open(fileStream);
            }
            catch
            {

            }
        }

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
    }
}
