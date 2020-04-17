using Microsoft.Win32;
using System;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using Button = System.Windows.Controls.Button;
using TextBox = System.Windows.Controls.TextBox;
using Window = System.Windows.Window;
using System.Diagnostics;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        String inputFilePath;
        public MainWindow()
        {
            InitializeComponent();
        }

        #region Browse file(s)
        private void Gomb_Click(object sender, RoutedEventArgs e)
        {
            Button senderButton = (Button)sender;
            String senderName = senderButton.Name;
            TextBox celmezo = null;

            celmezo = (senderName == "Inp") ? inputPath : outputPath;

            // Create an OpenFileDialog from Windows.Win32
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Set the filter according to the button pressed
            openFileDialog.Filter = 
                (senderName == "Inp") 
                ? "Excel Spreadsheet (*.xlsx)|*.xlsx|Text files (*.txt)|*.txt|All files (*.*)|*.*" 
                : "SEQ text (*.seq)|*.xlsx|Text files (*.txt)|*.txt|All files (*.*)|*.*";
            // openFileDialog.Filter = "ASCII text (*.asc)|*.xlsx|Text files (*.txt)|*.txt|All files (*.*)|*.*";
            
            //openFileDialog.ShowDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                celmezo.Text = openFileDialog.FileName;
                inputFilePath = openFileDialog.FileName;
                if (celmezo == inputPath) outputPath.Text = openFileDialog.FileName[0..^5] + "_SymbolTable.seq";
            }
        }
        #endregion


        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine(inputFilePath);
            Excel megnyitottFajl = new Excel(inputFilePath);
            string contents = ("\t" + "I 55.5" + "\t" + "Valtozonev" + "\t" + "comment" + "\n");
            string outputFileName = outputPath.Text;
            //if (!System.IO.File.Exists(outputFileName))
                System.IO.File.WriteAllText(outputFileName, contents);
        }

        class Excel
        {
            string path = "";
            _Application excel = new _Excel.Application();
            Workbook wb;
            _Worksheet ws;
           

            public Excel(string path)
            {
                this.path = path;
                wb = excel.Workbooks.Open(path);
                ws = (_Excel.Worksheet)excel.ActiveSheet;

                //ws = wb.Sheets[sheetNumber];
                _Excel.Range excelRange = ws.UsedRange;
                Debug.WriteLine(excelRange[2,2]);
                wb.Close(0);
            }
        }
    }
}
