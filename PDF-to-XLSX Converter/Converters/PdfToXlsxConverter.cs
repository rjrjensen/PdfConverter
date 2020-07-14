using System;
using System.Collections.Generic;
using System.Diagnostics;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.Office.Interop.Excel;
using Rectangle = iTextSharp.text.Rectangle;

namespace PDF_to_XLSX_Converter
{
    class PdfToXlsxConverter
    {
        private readonly PdfReader _pdfReader;
        private Application _application;
        private Workbook _workbook;
        private Worksheet _worksheet;

        private string _currentStartingCell = "A1";

        private Stopwatch _stopwatch;

        public PdfToXlsxConverter(string pdfFileLocation, string xlsxFileLocation)
        {
            _pdfReader = new PdfReader(pdfFileLocation);
            InitializeExcelWorksheet(xlsxFileLocation);
        }

        private void InitializeExcelWorksheet(string xlsxFileLocation)
        {
            _application = new Application {Visible = true};
            _workbook = _application.Workbooks.Add("");
            _worksheet = _workbook.ActiveSheet as Worksheet;
            _workbook.SaveAs(xlsxFileLocation, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }

        public void ExecuteOnColumnProfiles(List<PageRangeColumnProfile> pageRangeColumnProfiles)
        {
            foreach (var columnProfile in pageRangeColumnProfiles)
            {
                StartTimer();
                WritePageRangeToXlsxFile(columnProfile);
                StopTimer(columnProfile);
            }
            
            // _application.Visible = false;        TODO: REMOVE or USE
            // _application.UserControl = false;
            //
            // _workbook.Close();
            // _application.Quit();
        }

        private void WritePageRangeToXlsxFile(PageRangeColumnProfile columnProfile)
        {
            var startPage = columnProfile.StartPage;
            var endPage = columnProfile.EndPage;
            
            Console.WriteLine($"Starting range: {startPage} - {endPage}"); // TODO: REMOVE

            for (var page = startPage; page <= endPage; page++)
            {
                var textFromPage = GetTextFromPage(page, columnProfile);
                var cellRange = GetCellRange(textFromPage);
                
                WritePageToCorrespondingCells(textFromPage, cellRange);
                
                Save();
                
                ResetNewStartingCell(cellRange);

                Console.WriteLine($"Converted page: {page}"); // TODO: REMOVE
            }
        }

        private string[][] GetTextFromPage(int page, PageRangeColumnProfile columnProfile)
        {
            var numberOfColumns = columnProfile.Columns.Count;
            var textFromPage = new string[numberOfColumns][];
            
            for (var columnIndex = 0; columnIndex < numberOfColumns; columnIndex++)
            {
                var currentColumn = columnProfile.Columns[columnIndex];
                textFromPage[columnIndex] = ExtractCurrentColumnFromPage(page, currentColumn);
            }

            return textFromPage;
        }
        
        private string[] ExtractCurrentColumnFromPage(int page, Rectangle column)
        {
            var renderFilter = new RegionTextRenderFilter(column);
            var renderFilterArray = new RenderFilter[] {renderFilter};

            var filteredTextRenderListener = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), renderFilterArray);
            var textFromColumn = PdfTextExtractor.GetTextFromPage(_pdfReader, page, filteredTextRenderListener);

            return SplitColumnTextIntoRows(textFromColumn);
        }        
        
        private static string[] SplitColumnTextIntoRows(string textFromColumn) => textFromColumn.Split(new[] {Environment.NewLine}, StringSplitOptions.None);

        private CellRange GetCellRange(string[][] textFromPage)
        {
            return new CellRange()
            {
                StartingCell = _currentStartingCell,
                EndingCell = GetEndingCell(textFromPage)
            };
        }

        private string GetEndingCell(string[][] textFromPage)
        {
            var column = (char) (_currentStartingCell[0] + textFromPage.Length);
            var row = int.Parse(_currentStartingCell.Substring(1)) + textFromPage[0].Length;

            return $"{column}{row}";
        }

        private void WritePageToCorrespondingCells(string[][] textFromPage, CellRange cellRange)
        {
            _worksheet.Range[cellRange.StartingCell, cellRange.EndingCell].Value2 = textFromPage;
        }

        private void Save()
        {
            _workbook.Save();
        }

        private void ResetNewStartingCell(CellRange cellRange)
        {
            var nextRow = int.Parse(cellRange.EndingCell.Substring(1)) + 1;
            _currentStartingCell = $"A{nextRow}";
        }

        private void StartTimer()
        {
            _stopwatch = Stopwatch.StartNew();
        }

        private void StopTimer(PageRangeColumnProfile columnProfile)
        {
            _stopwatch.Stop();
            Console.WriteLine($"It took {_stopwatch.ElapsedMilliseconds} milliseconds to process pages {columnProfile.StartPage} - {columnProfile.EndPage}");
        }
    }
}