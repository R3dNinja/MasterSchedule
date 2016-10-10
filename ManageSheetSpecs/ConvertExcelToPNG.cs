using System;
using System.IO;
using System.Diagnostics;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using Syncfusion.XlsIO;

namespace ManageMasterSchedule
{
    public class ConvertExcelToPNG
    {
        public string sheetSize;
        public double formatHeight;
        public double centerLine;
        public int imagePerSheet;
        

        public int convertExceltoEMF(string path)
        {
            double sWidth = Command.thisCommand.getSheetWidth();
            setSheetParameters(sWidth);
            string sheetSize = Command.thisCommand.getSheetSize();

            if (sheetSize == "24 x 36")
            {
                formatHeight = 22.25;
                centerLine = 12.0;
                imagePerSheet = 4;
            }
            else if (sheetSize == "30 x 42")
            {
                formatHeight = 28.25;
                centerLine = 15.0;
                imagePerSheet = 5;
            }
            else if (sheetSize == "36 x 48")
            {
                formatHeight = 34.25;
                centerLine = 18.0;
                imagePerSheet = 6;
            }
            else
            {

            }

            double cellB;
            double cellC;
            double cellD;
            int resolution = 150;
            int pWidth;

            var savePath = Path.GetDirectoryName(path) + @"\Master Schedule (Images)\";
            if (Directory.Exists(savePath))
            {
                removeExistingImages(savePath);
            }
            else
            {
                Directory.CreateDirectory(savePath);
            }
            var saveName = "Master Schedule";

            Command.thisCommand.dialog.ConvertingWord();

            //Load the document.
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Excel2013;

            IWorkbook workbook = excelEngine.Excel.Workbooks.OpenReadOnly(path);
            IWorksheet sheet = workbook.Worksheets[1];

            //read widths of the cells we care about, B, C, and D
            cellB = sheet.GetColumnWidthInPixels(2);
            cellC = sheet.GetColumnWidthInPixels(3);
            cellD = sheet.GetColumnWidthInPixels(4);
            double cellWidth = cellB + cellC + cellD;

            float cellInches = Convert.ToSingle(cellWidth) / 72.0f;

            //store document width in pixels @ 150 DPI
            pWidth = Convert.ToInt32((cellInches * (float)resolution));

            double multiplier = 994.0 / pWidth;

            sheet.SetColumnWidthInPixels(2, Convert.ToInt32(cellB * multiplier));
            sheet.SetColumnWidthInPixels(3, Convert.ToInt32(cellC * multiplier));
            sheet.SetColumnWidthInPixels(4, Convert.ToInt32(cellD * multiplier));

            cellB = sheet.GetColumnWidthInPixels(2);
            cellC = sheet.GetColumnWidthInPixels(3);
            cellD = sheet.GetColumnWidthInPixels(4);
            cellWidth = cellB + cellC + cellD;

            cellInches = Convert.ToSingle(cellWidth) / 72.0f;

            //store document width in pixels @ 150 DPI
            pWidth = Convert.ToInt32((cellInches * (float)resolution));

            //Read the last used row in the doc
            int lastRow = sheet.UsedRange.LastRow;
            //Resize fonts in every cell
            for (int x = 1; x <= lastRow; x++)
            {
                string searchRow = "A" + x.ToString();
                double fntSize = sheet.Range[searchRow].CellStyle.Font.Size;
                var fgColor = sheet.Range[searchRow].CellStyle.PatternColorIndex;
                var bgColor = sheet.Range[searchRow].CellStyle.ColorIndex;
                ExcelKnownColors checkColor = (ExcelKnownColors)65;
                fixCellSize(sheet, x);

                /*if (fntSize != 18)
                {
                    if (bgColor == checkColor)
                    {
                        string searchRowB = "B" + x.ToString();
                        double fontSize = sheet.Range[searchRowB].CellStyle.Font.Size;
                        if (fontSize != 10)
                        {
                            fixCellSize(sheet, x);
                        }
                    }
                }*/
            }




            string fullRange = "B1:D" + lastRow.ToString();
            string lastRange = "D1:D" + lastRow.ToString();

            // do some setup on the sheet
            sheet.IsGridLinesVisible = false;
            sheet.Range[lastRange].WrapText = true;
            //sheet.Range[lastRange].AutofitRows();
            //sheet.Range[fullRange].AutofitRows();
            //sheet.Range[lastRange].AutofitRows();

            //delete hidden cells
            for (int rows = 1; rows <= lastRow; rows++)
            {
                bool row = sheet.IsRowVisible(rows);
                if (row == false)
                {
                    sheet.DeleteRow(rows);
                    rows--;
                }
            }

            lastRow = sheet.UsedRange.LastRow;



            //setup to figure out pages to publish
            double runningHeight = 0.0;

            ArrayList page = new ArrayList();
            ArrayList startCell = new ArrayList();
            ArrayList endCell = new ArrayList();

            int sPage = 1;
            int sCell = 1;

            //seperate rows into pages
            for (int rows = 1; rows <= lastRow; rows++)
            {

                string content = sheet.Range["B1"].Text;
                double tempRow1 = sheet.GetRowHeight(1);

                double tempRow = sheet.GetRowHeightInPixels(rows) / 72.0;
                runningHeight = runningHeight + tempRow;
                if (runningHeight > formatHeight)
                {
                    page.Add(sPage);
                    startCell.Add(sCell);
                    endCell.Add(rows - 1);
                    sPage++;
                    sCell = rows;
                    runningHeight = tempRow;
                }
                if (rows == lastRow)
                {
                    page.Add(sPage);
                    startCell.Add(sCell);
                    endCell.Add(rows);
                }
            }


            Command.thisCommand.dialog.setPageCount(page.Count);
            Command.thisCommand.dialog.SetupProgress(page.Count, "Task: Converting Excel Document to Images");


            //save document by row ranges as images
            for (int i = 0; i <= page.Count - 1; i++)
            {
                int pageValue = i + 1;
                runningHeight = 0.0;
                int startingCell = int.Parse(startCell[i].ToString());
                int endingCell = int.Parse(endCell[i].ToString());

                for (int x = startingCell; x < endingCell; x++)
                {
                    double tempRow = sheet.GetRowHeightInPixels(x) / 72.0;
                    runningHeight = runningHeight + tempRow;
                }
                float pageHeight = Convert.ToSingle(runningHeight);
                pageHeight = pageHeight * 150;


                Image image = sheet.ConvertToImage(startingCell, 2, endingCell, 4, ImageType.Metafile, null);
                Bitmap bitmap = null;

                bitmap = new Bitmap(pWidth, Convert.ToInt32(pageHeight));
                bitmap.SetResolution((float)resolution, (float)resolution);
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.High;
                    g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                    g.DrawImage(image, 0, 0, pWidth, Convert.ToInt32(pageHeight));
                    g.Dispose();
                }
                bitmap.Save(savePath + saveName + pageValue.ToString("D2") + ".png", ImageFormat.Png);
                bitmap.Dispose();

                //setup progress for export
                //int percentage = (int)Math.Round(((double)i / (double)page.Count) * 100.0);
                //if (OnProgressUpdate != null)
                //{
                //    OnProgressUpdate(Convert.ToInt32(percentage));
                //}
                //i++;
                Command.thisCommand.dialog.IncrementProgress();
                bitmap = null;
            }
            //close document without saving
            workbook.Close();
            excelEngine.ThrowNotSavedOnDestroy = false;
            excelEngine.Dispose();

            return page.Count;
        }

        private void fixCellSize(IWorksheet sheet, int x)
        {
            double initialRowHeight = sheet.GetRowHeight(x);

            string searchRowB = "B" + x.ToString();
            string searchRowC = "C" + x.ToString();
            string searchRowD = "D" + x.ToString();
            double tempRowHeight = sheet.GetRowHeight(x);
            //double foundFontSize = sheet.Range[searchRowB].CellStyle.Font.Size;
            //double newFontSize = (foundFontSize * multiplier);
            //double newFontSize = (foundFontSize - 1.75);
            //sheet.Range[searchRow].CellStyle.Font.Size = newFontSize;
            //sheet.Range[searchRowB].WrapText = true;
            sheet.Range[searchRowC].WrapText = true;
            sheet.Range[searchRowD].WrapText = true;
            sheet.Range[searchRowC].AutofitRows();
            double RowHeightC = sheet.GetRowHeight(x);
            //sheet.AutofitRow(x);
            sheet.Range[searchRowD].AutofitRows();
            double RowHeightD = sheet.GetRowHeight(x);
            if (RowHeightC > RowHeightD)
            {
                sheet.Range[searchRowC].AutofitRows();
            }
            //sheet.Range[searchRow].AutofitRows();

            //double tempRowHeight = sheet.GetRowHeight(x);

            /*if (tempRowHeight > initialRowHeight)
            {
                initialRowHeight = tempRowHeight;
            }

            searchRow = "C" + x.ToString();
            foundFontSize = sheet.Range[searchRow].CellStyle.Font.Size;
            //double newFontSize = (foundFontSize * multiplier);
            //double newFontSize = (foundFontSize - 1.75);
            //sheet.Range[searchRow].CellStyle.Font.Size = newFontSize;
            sheet.Range[searchRow].WrapText = true;
            sheet.Range[searchRow].AutofitRows();

            if (tempRowHeight > initialRowHeight)
            {
                initialRowHeight = tempRowHeight;
            }

            searchRow = "D" + x.ToString();
            foundFontSize = sheet.Range[searchRow].CellStyle.Font.Size;
            //double newFontSize = (foundFontSize * multiplier);
            //double newFontSize = (foundFontSize - 1.75);
            //sheet.Range[searchRow].CellStyle.Font.Size = newFontSize;
            sheet.Range[searchRow].WrapText = true;
            sheet.Range[searchRow].AutofitRows();

            if (tempRowHeight > initialRowHeight)
            {
                initialRowHeight = tempRowHeight;
            }

            sheet.Range[searchRow].RowHeight = initialRowHeight;*/
        }

        private void removeExistingImages(string savePath)
        {
            DirectoryInfo X = new DirectoryInfo(savePath);
            FileInfo[] listOfFiles = X.GetFiles("*.png");
            string[] Collection = new string[listOfFiles.Length];

            foreach (FileInfo FI in listOfFiles)
            {
                string fileToDelete = savePath + FI.Name;
                File.Delete(fileToDelete);
                //string sourceFileName = Source + FI.Name;
                //string destFileName = Destination + FI.Name;
                //File.Copy(sourceFileName, destFileName, true);
            }
        }

        private void setSheetParameters(double sWidth)
        {
            double swidth = sWidth;
            if (swidth == 3.0)
            {
                sheetSize = "24 x 36";
                Command.thisCommand.setSheetSize(sheetSize);
            }
            else if (swidth == 3.5)
            {
                sheetSize = "30 x 42";
                Command.thisCommand.setSheetSize(sheetSize);
            }
            else if (swidth == 4.0)
            {
                sheetSize = "36 x 48";
                Command.thisCommand.setSheetSize(sheetSize);
            }
            else
            {
                sheetSize = "void";
                Command.thisCommand.setSheetSize(sheetSize);
            }
        }
    }
}
