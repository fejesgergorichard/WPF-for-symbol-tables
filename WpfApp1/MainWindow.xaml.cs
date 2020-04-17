using Microsoft.Win32;
using System;
using System.Windows;
using System.Windows.Controls;
//using Microsoft.Office.Interop.Excel;
//using _Excel = Microsoft.Office.Interop.Excel;
using Button = System.Windows.Controls.Button;
using TextBox = System.Windows.Controls.TextBox;
using Window = System.Windows.Window;
using OfficeOpenXml;
using System.Text.RegularExpressions;
using System.IO;
using System.Globalization;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string inputFilePath;
        public MainWindow()
        {
            InitializeComponent();
        }

        #region Browse file(s)
        private void Gomb_Click(object sender, RoutedEventArgs e)
        {
            Button senderButton = (Button)sender;
            String senderName = senderButton.Name;


            // Check which Browse button was pressed
            TextBox celmezo = (senderName == "Inp") ? inputPath : outputPath;

            // Create an OpenFileDialog from Windows.Win32
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Set the filter according to the button pressed
            openFileDialog.Filter =
                (senderName == "Inp")
                ? "Excel Spreadsheet (*.xlsx)|*.xlsx|Text file (*.txt)|*.txt|All files (*.*)|*.*"
                : "SEQ text (*.seq)|*.seq|All files (*.*)|*.*";
                
            // If the dialog is shown, fill the TextBoxes with the paths for the files
            if (openFileDialog.ShowDialog() == true)
            {
                celmezo.Text = openFileDialog.FileName;
                inputFilePath = openFileDialog.FileName;
                if (celmezo == inputPath)
                {
                    outputPath.Text = openFileDialog.FileName.Substring(0, openFileDialog.FileName.Length - 5) + "_SymbolTable.seq";
                    outputSIMITPath.Text = openFileDialog.FileName.Substring(0, openFileDialog.FileName.Length - 5) + "_SIMIT_SHM.txt";
                }

            }
        }
        #endregion


        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //// 1       2       3       4       5       6       7               8                                              ///
            //// Symbol	InOut	Address	Type	Comment	Default	ImplicitSource	ImplicitSignal                                  ///
            //// masodik sor a fejlec, 3-tól adatok                                                                             ///
            //// Ami nekünk kell: \t  InOut(2) Address(3) \t Symbol(1) \t Comment(5) \n                                         ///
            //// Egyes mezők karakterhossza:                                                                                    ///
            ////                - szenzor:  14 karakter     PS_FCE_XA4_540          --> FCE_XA4.540                             ///
            ////                - MTR_RUN:  18 karakter     PS_MTR_RUN_XA4_510      --> MTR_RUN_XA4.510                         ///
            ////                - READY:    21 karakter     PS_READY_XA4_510_M100   --> READY_XA4.010_M100                      ///
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //string contents = ("\t" + "I 55.5" + "\t" + "Valtozonev" + "\t" + "comment" + "\n");

            /*
            //Creates a blank workbook. Use the using statment, so the package is disposed when we are done.
	        using (var p = new ExcelPackage())
            {
                //A workbook must have at least on cell, so lets add one... 
                var ws=p.Workbook.Worksheets.Add("MySheet");
                //To set values in the spreadsheet use the Cells indexer.
                ws.Cells["A1"].Value = "This is cell A1";
                //Save the new workbook. We haven't specified the filename so use the Save as method.
                p.SaveAs(new FileInfo(@"c:\workbooks\myworkbook.xlsx"));
           }
           */

            
            inputFilePath = inputPath.Text;

            // Create regex for checking the file extension
            Regex regexSeq = new Regex(@".seq");
            Regex regexText = new Regex(@".txt");
            
            if (regexSeq.IsMatch(inputFilePath))
                {

                // We need the '@' sign before the file path, because it's an unformatted string with special characters: '\'
                var megnyitottFajlInfo = new FileInfo(@inputFilePath);

                // Store the opened .xlsx file in variable 'p'
                var p = new ExcelPackage(megnyitottFajlInfo);

                // Open the sheet from the Excel file
                ExcelWorksheet ws = p.Workbook.Worksheets[1];
                
                // Get the length of the table
                int tableLength = ws.Dimension.End.Row;

                // Create or owerwrite the .seq and the .txt files with an empty string
                string outputFileName = outputPath.Text;
                string outputSIMITFileName = outputSIMITPath.Text;
                File.WriteAllText(outputFileName, "");
                File.WriteAllText(outputSIMITFileName, "");

                // j is the counter for the output files, since there could be empty rows in the Excel
                // It starts from the 3rd row (Excel indexing is one-based)
                int j = 3;

                // Iterate over the whole table
                String contentLine, inOut, address, symbol, comment;
                String SIMITLine;
                String[] SIMITLines = new string[tableLength];
                String[] contentLines = new string[tableLength];
                for (int i = 3; i <= tableLength; i++)
                {
                    symbol = (String)ws.Cells[i, 1].Value;
                    
                    // When the smybol is null or 5 spaces, it's an empty row
                    if (symbol == null || symbol.Equals("     "))
                        continue;

                    // Get rid of the 'PS_' part of the symbol names
                    symbol = symbol.correctSymbolName((bool)checkBoxRemovePS.IsChecked, (bool)checkBoxRepairDashes.IsChecked);

                    //inOut = (String)ws.Cells[i, 2].Value;
                    inOut = Ext.getInOut((String)ws.Cells[i, 2].Value, (bool)checkBoxFlipIQ.IsChecked);

                    address = (string)ws.Cells[i, 3].Value;

                    // Decrease the address of the outputs with the entered value
                    ////address = (float.Parse(address, CultureInfo.InvariantCulture.NumberFormat) + 2).ToString(CultureInfo.InvariantCulture);
                    if (inOut.Equals("Q"))
                        address = decreaseOutputAddress(address, differenceI_Q.Text);

                    
                    comment = (String)ws.Cells[i, 5].Value;
                    
                    // Create a line of the .seq file
                    contentLine = "\t" + inOut + " " + address + "\t" + symbol + "\t" + comment;

                    // Create a line of the SIMIT .txt file
                    SIMITLine = "";
                    for (int k = 1; k <= 8; k++)
                    {
                        // If every signal is supposed to be mapped
                        if (!(bool)((CheckBox)mapEverySignal).IsChecked)
                        {
                            // Add every cell except the last one
                            if (k != 8)
                            {
                                SIMITLine += (string)ws.Cells[i, k].Value;
                                SIMITLine += "\t";
                            }

                            // If the last cell is empty, add it as it is
                            if ((((string)ws.Cells[i, k].Value).Equals("     ")) && k == 8)
                            {
                                SIMITLine += (string)ws.Cells[i, k].Value;
                            }

                            // If the last cell is not empty (it's a mapped signal) then add the corrected symbol name to the table
                            else if (!(((string)ws.Cells[i, k].Value).Equals("     ")) && k == 8)
                            {
                                SIMITLine += symbol;
                            }
                        }
                        else
                        {
                            if (k != 7 && k != 8)
                            {
                                SIMITLine += (string)ws.Cells[i, k].Value;
                                SIMITLine += "\t";
                            }
                            else
                            {
                                SIMITLine += PLCName.Text + "\t" + symbol;
                                k++;
                            }
                        }
                    }
                

                    // Add the created .seq line to the contentLines array
                    contentLines[j - 3] = contentLine;

                    // Add the created SIMIT line to the SIMITLines array
                    SIMITLines[j - 1] = SIMITLine;

                    // Next output line
                    j++;

                }

                // Header lines for the SIMIT .txt
                SIMITLines[0] = "#Signal properties; SIMIT V9.0; COUPLING:proba1";
                SIMITLines[1] = "Symbol\tInOut\tAddress\tType\tComment\tDefault\tImplicitSource\tImplicitSignal";

                // Write the contentLines array in the output .seq file
                File.WriteAllLines(outputFileName, contentLines);

                // Write the SIMITLines array in the output .txt file
                File.WriteAllLines(outputSIMITFileName, SIMITLines);

                // Close the Excel file
                p.Dispose();

                MessageBox.Show("Conversion done!\nCongratulations!\nEnjoy programming your Siemens PLC :)", "Success congratulations happy new year merry christmas dorime ameno dorime ameno ameno latire latiremo");
            }

            if (!regexSeq.IsMatch(inputFilePath))
            {
                MessageBox.Show("Please browse for a .seq Symbol Table file!");
            }
        }

        private void differenceI_Q_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            // Create regex for checking the range  of 0-9
            Regex regexNum = new Regex("[^0-9]");
            
            // The textBox is only handled, when the entered text matches the above regular expression
            e.Handled = regexNum.IsMatch(e.Text);
        }

        private string decreaseOutputAddress(string originalAddress, string differenceI_Q)
        {
            string newAddress;
            int difference = Int32.Parse(differenceI_Q);
            
            // Culture info is needed for the decimal points and commas
            newAddress= String.Format(CultureInfo.InvariantCulture, "{0:0.0}", (double.Parse(originalAddress, CultureInfo.InvariantCulture) - difference));
            
            return newAddress;
        }

        

        private void dummyMethod()
        {
            if ((bool)checkBoxRemovePS.IsChecked)
                pityu.Content = "checked";
            if (!(bool)checkBoxRemovePS.IsChecked)
                pityu.Content = "not checked";
        }

        // Enables the mapping of every signal to the PLC
        private void mapEverySignal_Click(object sender, RoutedEventArgs e)
        {
            PLCName.IsEnabled = (bool)((CheckBox)sender).IsChecked;
        }


        /*
        class Excel
        {
            string path = "";
            _Application excel = new _Excel.Application();
            Workbook wb;
            _Worksheet ws;
            public _Excel.Range excelRange;
            public string cont;


            public Excel(string path)
            {
                this.path = path;
                wb = excel.Workbooks.Open(path);
                ws = (_Excel.Worksheet)excel.Sheets[1];

                //ws = wb.Sheets[sheetNumber];
                this.excelRange = ws.UsedRange;
                //Debug.WriteLine(excelRange[2, 2]);
                cont = (string)excelRange[2, 2].Value;
                
                wb.Close(0);
            }

        }
        */
        }
    }

// Extension class for functions
public static class Ext
{
    // Function for replacing a single character of the string at a given index
    public static string ReplaceAt(this string input, int index, char newChar)
    {
        if (input == null)
            throw new ArgumentNullException("input");
        char[] chars = input.ToCharArray();
        chars[index] = newChar;

        return new string(chars);
    }


    /// Function for removing the PS_ from the beginning of symbols and correcting the underscores to dots
    public static string correctSymbolName(this string originalSymbol, bool removePS_, bool repairDashes)
    {
        string newSymbol;
        if (removePS_)
            newSymbol = originalSymbol.Substring(3);
        else
            newSymbol = originalSymbol;

        if (repairDashes)
        {
            if (// Inputs
            newSymbol.Substring(0, 3) == "FCE" ||
            newSymbol.Substring(0, 3) == "FCI" ||
            newSymbol.Substring(0, 3) == "ENC" ||
            newSymbol.Substring(0, 7) == "MTR_RUN" ||
            // Output
            newSymbol.Substring(0, 3) == "OTR" ||
            newSymbol.Substring(0, 3) == "DQS" ||
            newSymbol.Substring(0, 4) == "OTRR"
            )
                newSymbol = newSymbol.ReplaceAt(newSymbol.Length - 4, '.');

            else if (newSymbol.Substring(0, 5) == "READY")
                newSymbol = newSymbol.ReplaceAt(newSymbol.Length - 9, '.');
        }

        if (newSymbol.Length > 24)
        {
            Console.WriteLine("Symbol longer than 24 characters: " + newSymbol + "  Length: " + newSymbol.Length);
            newSymbol = newSymbol.Substring(0, 24);
        }

        return newSymbol;
    }

    /// Function for getting the correct I/Q addresses. They need to be the exact opposite of the I/Q-s in Plant Simulation or SIMIT
    public static string getInOut(string inOutFromExcel, bool flipIQ)
    {
        string correctedInOut = " ";

        if (flipIQ)
        {
            if (inOutFromExcel.Equals("Q"))
                correctedInOut = "I";
            if (inOutFromExcel.Equals("I"))
                correctedInOut = "Q";
        }    
        else
            correctedInOut = inOutFromExcel;


        return correctedInOut;
    }
}
