using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2WikiTable
{
    enum emCellStatus { Nonne, UniqueCell, MergeCellHeader, MergeCellBody };

    public partial class Excel2WikiTable : Form
    {

        object missingValue = System.Reflection.Missing.Value;
        Excel.Application excelApp = null;
        float dpiX = 0;
        float dpiY = 0;

        public Excel2WikiTable()
        {
            InitializeComponent();
            Graphics graphics = this.CreateGraphics();
            dpiX = graphics.DpiX;
            dpiY = graphics.DpiY;
        }

        private bool MakeHtmlTable(string outputFileName, List<List<SimpleExcelCell>> processDataTable)
        {
            try
            {
                using (StreamWriter outputFile = new StreamWriter(outputFileName))
                {
                    outputFile.WriteLine("<table>");
                    foreach (List<SimpleExcelCell> row in processDataTable)
                    {
                        outputFile.Write("\t<tr>");
                        foreach (SimpleExcelCell column in row)
                        {
                            System.Windows.Forms.Application.DoEvents();
                            outputFile.Write("<td width=\"" + column.CellWidthPixels + "\">");
                            outputFile.Write(column.CellValue);
                            outputFile.Write("</td>");
                        }
                        outputFile.WriteLine("</tr>");
                    }
                    outputFile.WriteLine("</table>");
                    outputFile.Flush();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
            return true;
        }

        private List<List<SimpleExcelCell>> ProcessWorkSheet(Worksheet worksheet)
        {

            int usedColumnsCount = worksheet.UsedRange.Cells.Columns.Count; // Total columns count
            int usedRowsCount = worksheet.UsedRange.Cells.Rows.Count; // Total rows count
            int[] columnWidrhPixel = new int[usedColumnsCount];

            #region Get every column width and save it in array.
            for (int i = 1; i <= usedRowsCount; i++)
            {
                int thisRowColumnsCount = worksheet.UsedRange.Rows[i].Columns.Count;
                if (thisRowColumnsCount < usedColumnsCount) // If thisRowColumnsCount less then usedColumnsCount, this row has merge column
                {
                    continue;
                }
                for (int j = 1; j <= thisRowColumnsCount; j++)
                {
                    columnWidrhPixel[j - 1] = Points2Pixels(worksheet.Cells[i, j].Width);
                }
                break;  // If get all columns width, break out this loop
            }
            #endregion

            #region Get excel value
            List<List<SimpleExcelCell>> rowList = new List<List<SimpleExcelCell>>();
            for (int i = 1; i <= usedRowsCount; i++)
            {
                List<SimpleExcelCell> columnList = new List<SimpleExcelCell>();
                int columnsCount = worksheet.UsedRange.Rows[i].Columns.Count;
                int widthTailIndex = usedColumnsCount;
                int currentCellIndex = 0;

                for (int j = 1; j <= columnsCount; j++)
                {
                    System.Windows.Forms.Application.DoEvents();
                    // Check excel range(cell) merge and value status.
                    emCellStatus cellStatus = ConvertCellStatus(worksheet.Cells[i, j], j);
                    if (cellStatus == emCellStatus.UniqueCell || cellStatus == emCellStatus.MergeCellHeader)
                    {
                        string currentCellName = worksheet.Cells[i, j].Address.ToString();
                        string currentCellValue = worksheet.Cells[i, j].Value2 == null ? "" : worksheet.Cells[i, j].Value2.ToString();
                        columnList.Add(new SimpleExcelCell(currentCellName, currentCellValue, cellStatus));
                        columnList[currentCellIndex].CellWidthPixels = columnWidrhPixel[j - 1];
                        currentCellIndex++;
                    }
                    else if (cellStatus == emCellStatus.MergeCellBody)
                    {
                        columnList[currentCellIndex - 1].CellWidthPixels += columnWidrhPixel[j - 1];
                        continue;
                    }
                    else
                    {
                        throw new Exception("Error:emCellStatus = " + cellStatus.ToString());
                    }
                }
                rowList.Add(columnList);
            }
            #endregion
            return rowList;
        }

        private Worksheet GetExcelWorkSheet(string excelFileName, int workSheetNumber)
        {
            try
            {
                excelApp = new Excel.Application();
                if (excelApp == null)
                {
                    throw new ArgumentNullException("excelApp Open Failed");
                }
                else
                {
                    excelApp.Visible = false;
                    excelApp.UserControl = true;
                    Workbook excelWorkBook = excelApp.Application.Workbooks.Open(excelFileName, missingValue, true, missingValue,
                        missingValue, missingValue, missingValue, missingValue, missingValue, true, missingValue,
                        missingValue, missingValue, missingValue, missingValue);
                    return (Worksheet)excelWorkBook.Worksheets.get_Item(workSheetNumber);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                excelApp.Quit();
                excelApp = null;
                return null;
            }
        }

        private void KillExcelProcess()
        {
            Process[] procs = Process.GetProcessesByName("excel");
            foreach (Process pro in procs)
            {
                pro.Kill(); //Kill process is a easy way to do
            }
        }

        private void btnParse_Click(object sender, EventArgs e)
        {

            string[] exceFilesList = StackBasedIteration.TraverseTree(textBox1.Text);


            for (int i = 0; i < exceFilesList.Length; i++)
            {
                string excelFile = exceFilesList[i];
                if (File.Exists(System.IO.Path.ChangeExtension(excelFile, "md")))
                {
                    textBox2.AppendText(System.IO.Path.ChangeExtension(excelFile, "md") + "...already exist.\r\n");
                }
                Worksheet worksheet = GetExcelWorkSheet(excelFile, 1);
                List<List<SimpleExcelCell>> processDataTable = ProcessWorkSheet(worksheet);
                MakeHtmlTable(System.IO.Path.ChangeExtension(excelFile, "md"), processDataTable);
                textBox2.AppendText(System.IO.Path.ChangeExtension(excelFile, "md") + "...OK!\r\n");
                if (i % 10 == 0)
                {
                    KillExcelProcess();
                }
            }
            MessageBox.Show("Done!");

        }

        //private int Alphabet2Number(string strAlphabet)
        //{
        //    string strNumber = string.Empty;
        //    foreach (char charAlphabet in strAlphabet)
        //    {
        //        int index = (char.ToUpper(charAlphabet) - 'A') + 1; // Alphabet index start from 1
        //        strNumber += index;
        //    }
        //    return Convert.ToInt32(strNumber);
        //}

        private int Points2Pixels(double iPoint)
        {
            return Convert.ToInt32(iPoint * dpiX / 72);
        }

        private emCellStatus ConvertCellStatus(Range range, int xAddress)
        {
            bool bMergeCell = range.MergeCells;
            bool bHasValue = range.Value2 == null ? false : true;
            if (bMergeCell)
            {
                if (bHasValue)
                {
                    return emCellStatus.MergeCellHeader;
                }
                else
                {
                    if (xAddress == 1)  // If address x is the first column, it should be cell header
                    {
                        return emCellStatus.MergeCellHeader;
                    }
                    else
                    {
                        return emCellStatus.MergeCellBody;
                    }
                }
            }
            else
            {
                return emCellStatus.UniqueCell;
            }
        }

        private void btnSelectDirectory_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void Excel2WikiTable_FormClosed(object sender, FormClosedEventArgs e)
        {
            KillExcelProcess();
        }
    }

    class SimpleExcelCell   // Make simple data struct, let us read data clearly.
    {
        public string CellName = string.Empty;
        public string CellValue = string.Empty;
        public int CellWidthPixels = 0;
        public emCellStatus CellStatus = emCellStatus.Nonne;

        public SimpleExcelCell(string strName, string strValue, emCellStatus cellStatus)
        {
            this.CellName = strName;
            this.CellValue = strValue;
            this.CellStatus = cellStatus;
        }
    }

    class StackBasedIteration
    {
        public static string[] TraverseTree(string root)
        {
            List<string> excelFilesList = new List<string>();
            excelFilesList.Clear();

            // Data structure to hold names of subfolders to be
            // examined for files.
            Stack<string> dirs = new Stack<string>(20);

            if (!System.IO.Directory.Exists(root))
            {
                throw new ArgumentException();
            }
            dirs.Push(root);

            while (dirs.Count > 0)
            {
                string currentDir = dirs.Pop();
                string[] subDirs;
                try
                {
                    subDirs = System.IO.Directory.GetDirectories(currentDir);
                }
                // An UnauthorizedAccessException exception will be thrown if we do not have
                // discovery permission on a folder or file. It may or may not be acceptable 
                // to ignore the exception and continue enumerating the remaining files and 
                // folders. It is also possible (but unlikely) that a DirectoryNotFound exception 
                // will be raised. This will happen if currentDir has been deleted by
                // another application or thread after our call to Directory.Exists. The 
                // choice of which exceptions to catch depends entirely on the specific task 
                // you are intending to perform and also on how much you know with certainty 
                // about the systems on which this code will run.
                catch (UnauthorizedAccessException e)
                {
                    Console.WriteLine(e.Message);
                    continue;
                }
                catch (System.IO.DirectoryNotFoundException e)
                {
                    Console.WriteLine(e.Message);
                    continue;
                }

                string[] files = null;
                try
                {
                    files = System.IO.Directory.GetFiles(currentDir, "*.xls");
                    excelFilesList.AddRange(files);
                }
                catch (UnauthorizedAccessException e)
                {

                    Console.WriteLine(e.Message);
                    continue;
                }
                catch (System.IO.DirectoryNotFoundException e)
                {
                    Console.WriteLine(e.Message);
                    continue;
                }
                // Push the subdirectories onto the stack for traversal.
                // This could also be done before handing the files.
                foreach (string str in subDirs)
                    dirs.Push(str);
            }
            return excelFilesList.ToArray();
        }
    }
}
