using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Windows.Forms;



namespace LevenshteinCalculations
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window

    {
        int SumOfWords;
        int SumOfDistances;



        List<WordPair> currentPairs = new List<WordPair>();
        Calculator calculator = new Calculator();
        public MainWindow()
        {
            InitializeComponent();
        }

        private void TargetString_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            var textBox = sender as System.Windows.Controls.TextBox;

            if (textBox != null && !textBox.IsReadOnly && e.KeyboardDevice.IsKeyDown(Key.Tab))
            {
                textBox.SelectAll();
            }
            
                
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (TargetString.Text.Length == 0)
            {
                System.Windows.MessageBox.Show("Target cannot be empty");
                return;
            }
            if (SourceText.Text.Length == 0)
            {
                System.Windows.MessageBox.Show("Source cannot be empty");
                return;
            }
            if (TotalWeight.Text.Length == 0)
            {
                System.Windows.MessageBox.Show("Weight cannot be empty");
                return;
            }
            String excludeWord = System.Configuration.ConfigurationManager.AppSettings["ExcludedWords"];
            excludeWord = excludeWord.ToLower();
            String[] excludedWords = excludeWord.Split(new char[] { ',' });
            string Target = TargetString.Text;
            string Source = SourceText.Text;
            string[] sourcesplit = Source.Split(new char[] { ' ', '/', ',', '-', '.', '+', '?', '&', '\'', '\"', '!', '*', '_' });
            string[] targetsplit = Target.Split(new char[] { ' ', '/', ',', '-', '.', '+', '?', '&', '\'', '\"', '!', '*', '_' });
            WordPair[] wordPairs = new WordPair[targetsplit.Length * sourcesplit.Length];
            int k = 0;
            for (int i = 0; i < sourcesplit.Length; i++)
            {

                for (int j = 0; j < targetsplit.Length; j++)
                {
                    WordPair temppair = new WordPair();
                    temppair.SourceWord = sourcesplit[i];
                    temppair.SourceID = i;
                    temppair.TargetID = j;
                    temppair.TargetWord = targetsplit[j];
                    for (int k2 = 0; k2 < excludedWords.Length; k2++)
                    {
                        if (temppair.SourceWord.ToLower() == excludedWords[k2].Trim().ToLower() || temppair.TargetWord.ToLower() == excludedWords[k2].Trim())
                        {
                            temppair.excluded = true;

                        }
                        if (chkTwoWord.IsChecked == false) { 
                            if (temppair.SourceWord.Length < 3 || temppair.TargetWord.Length < 3)
                            {
                                temppair.excluded = true;
                            }
                        
                        }


                    }


                    wordPairs[k++] = temppair;


                }

            }
            wordPairs = calculator.CalcLevDistance(wordPairs);
            wordPairs = calculator.CalcDistances(wordPairs);
            wordPairs = calculator.CalcIndividualScore(wordPairs);
            wordPairs = calculator.FindScored(wordPairs, Math.Max(targetsplit.Length, sourcesplit.Length));
            var totalScore = calculator.FindTotalScore(wordPairs);
            int reduction = 0;
            int reductionBy = 0;
            try
            {
                reductionBy = Convert.ToInt32(txtSingleWord.Text);
            }
            catch
            {
                reductionBy = 0;
                txtSingleWord.Text = "Invalid entry, reduction set to 0.";
            }
            if (totalScore.Item1 == 100 && chkSingleWord.IsChecked == true && sourcesplit.Count() < )
            {
                if (sourcesplit.Length == 1)
                {
                    totalScore.Item1 = totalScore.Item1 - ((targetsplit.Length - 1) * Convert.ToInt32(txtSingleWord.Text));
                    reduction = (targetsplit.Length - 1) * reductionBy;
                }
                if (targetsplit.Length == 1)
                {

                    {
                        totalScore.Item1 = totalScore.Item1 - ((sourcesplit.Length - 1) * Convert.ToInt32(txtSingleWord.Text));
                        reduction = (sourcesplit.Length - 1) * Convert.ToInt32(txtSingleWord.Text);
                    }

            }
            if (reduction > 0 || totalScore.Item1 == 100)
            {
                totalScore.Item1 = totalScore.Item1 - 1;

            }

            foreach (WordPair pair in wordPairs)
            {
                pair.InitialScore = decimal.Round(pair.InitialScore, 3);
            }
            WordGrid.ItemsSource = wordPairs;

            WithoutWeight.Content = totalScore.Item1.ToString("#.##");
            if (totalScore.Item1 == 0)
            {
                WithoutWeight.Content = "0";
            }

            if (TotalWeight.Text == "Weight")
            {
                TotalWeight.Text = "100";
            }
            SumOfWords = totalScore.Item2;
            SumOfDistances = totalScore.Item3;
            lblSumOfWords.Content = SumOfWords.ToString();
            lblSumOfDistances.Content = SumOfDistances.ToString();

            lblTotalScoreCalc.Content = $"(({SumOfWords}-{SumOfDistances})*100)/{SumOfWords}={totalScore.Item1}";
            if (reduction > 0)
            {
                lblTotalScoreCalc.Content = $"(({SumOfWords}-{SumOfDistances})*100)/{SumOfWords} - {reduction}={totalScore.Item1}";
            }
            totalScore.Item1 = totalScore.Item1 * (decimal.Parse(TotalWeight.Text) / 100);
            WithWeight.Content = totalScore.Item1.ToString("#.##");
            if (totalScore.Item1 == 0)
            {
                WithWeight.Content = "0";
            }
            currentPairs = wordPairs.ToList();

            

            

        }

        private void TextBox_TextChanged_1(object sender, TextChangedEventArgs e)
        {

        }

        private void SourceText_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void WordGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void RowClicked(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var currentRowindex = WordGrid.Items.IndexOf(WordGrid.SelectedItem);
            {
                if (WordGrid.SelectedItem != null)
                {
                    WordPair passPair = new WordPair();
                    passPair = (WordPair)WordGrid.Items.GetItemAt(currentRowindex);
                    CalcWindow cal = new CalcWindow(passPair);
                    cal.Show();
                }
            }



        }

        private void toExcelButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                using var dialog = new FolderBrowserDialog
                {
                    Description = "Time to select a folder",
                    UseDescriptionForTitle = true,
                    SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                    + Path.DirectorySeparatorChar,
                    ShowNewFolderButton = true
                };
                List<WordPair> list = new List<WordPair>();
                WordPair[] wordPairs = new WordPair[WordGrid.Items.Count];
                int count2 = 0;
                foreach (WordPair wordPair in WordGrid.ItemsSource)
                {
                    wordPairs[count2] = wordPair;
                    count2++;
                }
                list = wordPairs.ToList();
                Excel.Application excapp = new Microsoft.Office.Interop.Excel.Application();



                var workbook = excapp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

                var sheet = (Excel.Worksheet)workbook.Worksheets[1];



                string cellName;
                int counter = 2;
                cellName = "A1";
                var rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "SourceWord";
                cellName = "B1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "TargetWord";
                cellName = "C1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "Scored";
                cellName = "D1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "Excluded";
                cellName = "E1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "LevenshteinDistance";
                cellName = "F1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "SourceDistance";
                cellName = "G1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "TargetDistance";
                cellName = "H1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "TotalDistance";
                cellName = "I1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "InitialScore";
                foreach (WordPair pair in list)
                {
                    cellName = "A" + counter.ToString();
                    var range = sheet.get_Range(cellName, cellName);
                    range.Value2 = pair.SourceWord.ToString();
                    cellName = "B" + counter.ToString();
                    range = sheet.get_Range(cellName, cellName);
                    range.Value2 = pair.TargetWord.ToString();
                    cellName = "C" + counter.ToString();
                    range = sheet.get_Range(cellName, cellName);
                    if (pair.scored == true)
                    {
                        range.Value2 = "Yes";
                    }
                    else
                    {
                        range.Value2 = "No";
                    }


                    cellName = "D" + counter.ToString();
                    range = sheet.get_Range(cellName, cellName);
                    if (pair.excluded == true)
                    {
                        range.Value2 = "Yes";
                    }
                    else
                    {
                        range.Value2 = "No";
                    }

                    cellName = "E" + counter.ToString();
                    range = sheet.get_Range(cellName, cellName);
                    range.Value2 = pair.levdistance.ToString();
                    cellName = "F" + counter.ToString();
                    range = sheet.get_Range(cellName, cellName);
                    range.Value2 = pair.sourcedistance.ToString();
                    cellName = "G" + counter.ToString();
                    range = sheet.get_Range(cellName, cellName);
                    range.Value2 = pair.targetdistance.ToString();
                    cellName = "H" + counter.ToString();
                    range = sheet.get_Range(cellName, cellName);
                    range.Value2 = pair.totaldistance.ToString();
                    cellName = "I" + counter.ToString();
                    range = sheet.get_Range(cellName, cellName);
                    range.Value2 = pair.InitialScore.ToString();
                    ++counter;


                }

                cellName = "K1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "Score Without Weight";
                cellName = "K2";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = WithoutWeight.Content.ToString();
                cellName = "L1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "Score With Weight";
                cellName = "L2";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = WithWeight.Content.ToString();
                cellName = "M1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "Sum of Words";
                cellName = "M2";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = lblSumOfWords.Content.ToString();
                cellName = "N1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "Sum of Distances";
                cellName = "N2";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = lblSumOfDistances.Content.ToString();
                cellName = "O1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "Total Score Calculation";
                cellName = "O2";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = lblTotalScoreCalc.Content.ToString();
                workbook.Save();




                excapp.Visible = true;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.ToString());
            }
            



        }

        private void exportButton_Click(object sender, RoutedEventArgs e)
        {
            var FolderDialog = new FolderBrowserDialog();
            FolderDialog.ShowDialog();
            if (fileName.Text.Length == 0)
            {
                System.Windows.MessageBox.Show("File Name cannot be empty");
                return;
            }

            try
            {

                using var dialog = new FolderBrowserDialog
                {
                    Description = "Time to select a folder",
                    UseDescriptionForTitle = true,
                    SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                    + Path.DirectorySeparatorChar,
                    ShowNewFolderButton = true
                };
                List<WordPair> list = new List<WordPair>();
                WordPair[] wordPairs = new WordPair[WordGrid.Items.Count];
                int count2 = 0;
                foreach (WordPair wordPair in WordGrid.ItemsSource)
                {
                    wordPairs[count2] = wordPair;
                    count2++;
                }
                list = wordPairs.ToList();
                Excel.Application excapp = new Microsoft.Office.Interop.Excel.Application();



                var workbook = excapp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

                var sheet = (Excel.Worksheet)workbook.Worksheets[1];



                string cellName;
                int counter = 2;
                cellName = "A1";
                var rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "SourceWord";
                cellName = "B1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "TargetWord";
                cellName = "C1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "Scored";
                cellName = "D1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "Excluded";
                cellName = "E1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "LevenshteinDistance";
                cellName = "F1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "SourceDistance";
                cellName = "G1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "TargetDistance";
                cellName = "H1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "TotalDistance";
                cellName = "I1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "InitialScore";
                foreach (WordPair pair in list)
                {
                    cellName = "A" + counter.ToString();
                    var range = sheet.get_Range(cellName, cellName);
                    range.Value2 = pair.SourceWord.ToString();
                    cellName = "B" + counter.ToString();
                    range = sheet.get_Range(cellName, cellName);
                    range.Value2 = pair.TargetWord.ToString();
                    cellName = "C" + counter.ToString();
                    range = sheet.get_Range(cellName, cellName);
                    if (pair.scored == true)
                    {
                        range.Value2 = "Yes";
                    }
                    else
                    {
                        range.Value2 = "No";
                    }


                    cellName = "D" + counter.ToString();
                    range = sheet.get_Range(cellName, cellName);
                    if (pair.excluded == true)
                    {
                        range.Value2 = "Yes";
                    }
                    else
                    {
                        range.Value2 = "No";
                    }

                    cellName = "E" + counter.ToString();
                    range = sheet.get_Range(cellName, cellName);
                    range.Value2 = pair.levdistance.ToString();
                    cellName = "F" + counter.ToString();
                    range = sheet.get_Range(cellName, cellName);
                    range.Value2 = pair.sourcedistance.ToString();
                    cellName = "G" + counter.ToString();
                    range = sheet.get_Range(cellName, cellName);
                    range.Value2 = pair.targetdistance.ToString();
                    cellName = "H" + counter.ToString();
                    range = sheet.get_Range(cellName, cellName);
                    range.Value2 = pair.totaldistance.ToString();
                    cellName = "I" + counter.ToString();
                    range = sheet.get_Range(cellName, cellName);
                    range.Value2 = pair.InitialScore.ToString();
                    ++counter;


                }

                cellName = "K1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "Score Without Weight";
                cellName = "K2";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = WithoutWeight.Content.ToString();
                cellName = "L1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "Score With Weight";
                cellName = "L2";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = WithWeight.Content.ToString();
                cellName = "M1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "Sum of Words";
                cellName = "M2";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = lblSumOfWords.Content.ToString();
                cellName = "N1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "Sum of Distances";
                cellName = "N2";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = lblSumOfDistances.Content.ToString();
                cellName = "O1";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = "Total Score Calculation";
                cellName = "O2";
                rangeH = sheet.get_Range(cellName, cellName);
                rangeH.Value = lblTotalScoreCalc.Content.ToString();





                workbook.SaveAs(FolderDialog.SelectedPath + @"\" +fileName.Text, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
               false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.ToString());
            }


        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
 
}