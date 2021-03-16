using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;

namespace Extreme_MSSQL_Checker
{
    public partial class MainWindow : Window
    {
        List<string> csv;

        List<(string server, string serverTrim, string serverType)> spotTrimList;
            public MainWindow()
            {
                InitializeComponent();
            }

            private void ExcelImport_Click(object sender, RoutedEventArgs e)
            {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files |*.xlsx"
            };
            if ((bool)!openFileDialog.ShowDialog()) return;

            spotTrimList = new List<(string server, string serverTrim, string serverType)>();
                var count = 0;

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBook;

                try
                {
                    xlWorkBook = xlApp.Workbooks.Open(openFileDialog.FileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 0);
                    
                    Excel.Worksheet worksheet1 = (Excel.Worksheet)xlApp.Worksheets["Sheet1"];

                    int sheet1LastRowCount = worksheet1.UsedRange.Rows.Count;

                    for (var i = 3; i <= sheet1LastRowCount; i++)
                    {

                        count++;

                        var input1 = worksheet1.Range["A" + i, "A" + i].Value2.ToString();
                        var input2 = worksheet1.Range["Q" + i, "Q" + i].Value2.ToString();

                    if (input1 == null || input2 == null) continue;

                            var index = input1.IndexOf("0", StringComparison.Ordinal);

                            if (index <= 0) continue;

                        var input1Trim = input1.Substring(0, index);

                            spotTrimList.Add((input1, input1Trim, input2));
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    xlApp.Quit();
                }

                ServerLabel.Content = $"Server caricati in memoria: {spotTrimList.Count} su un totale di: {count}";

                if(csv?.Count > 0) CompareExcelCsv();

        }

             private void CompareExcelCsv()
              {
                  OutputBox.Text = string.Empty;
                  if (spotTrimList.Count == 0 || csv?.Count == 0)
                  {
                      MessageBox.Show("CSV o Excel Non caricato o lista server non valida!", "Errore", MessageBoxButton.OK, MessageBoxImage.Error);
                      return;
                  }

                  var csv2 = csv.Where(s => spotTrimList.Find(x => s.Contains(x.server)).serverType == "DB-CL").ToList();
                  var spotTrimList2 = spotTrimList.Where(x => x.serverType == "DB-CL").ToList();

                  foreach (var s in spotTrimList2)
                  {
                      //MessageBox.Show(s.server);
                      if (csv2.Any(x => x.Contains(s.serverTrim))) continue;
                      OutputBox.Text += s.server + "\n";
                  }

                  if (OutputBox.Text == string.Empty) MessageBox.Show("Nessun server assente dalla lista!", "Attenzione!", MessageBoxButton.OK, MessageBoxImage.Information);
              }

            private void CSVImport_Click(object sender, RoutedEventArgs e)
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "CSV Files |*.csv";
                if ((bool)!openFileDialog.ShowDialog()) return;

                csv = new List<string>();

                var contents = Array.Empty<string>();

                try
                {
                    contents = File.ReadAllText(openFileDialog.FileName).Split('\n');
                }

                catch (IOException)
                {
                    MessageBox.Show("Eccezione I/O! Impossibile aprire il CSV, verificare che non sia aperto in un altro programma!", "Errore", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (contents.Length <= 1)
                {
                    MessageBox.Show("CSV non valido o vuoto!", "Errore", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                foreach (var s in contents)
                {
                    var input = s.Split(",")[2].Trim(new Char[] {'"'});

                    /*var index = input.IndexOf("0", StringComparison.Ordinal);
                    if (index > 0)
                        input = input.Substring(0, index);*/
                    csv.Add(input);
                }

                csv.RemoveAt(0);    //Rimuove deviceHostName

                if(spotTrimList?.Count > 0) CompareExcelCsv();
            }
        }
    }