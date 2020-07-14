using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Rectangle = iTextSharp.text.Rectangle;

namespace PDF_to_XLSX_Converter
{
    class PdfToCsvConverter
    {
        private string _pdfFileLocation;
        private string _csvFileLocation;

        public PdfToCsvConverter(string pdfFileLocation, string csvFileLocation)
        {
            _pdfFileLocation = pdfFileLocation;
            _csvFileLocation = csvFileLocation;
        }

        public void ExecuteOnColumnProfiles(List<PageRangeColumnProfile> pageRangeColumnProfiles)
        {
            using var reader = new PdfReader(_pdfFileLocation);
            using var fileStream = new FileStream(_csvFileLocation, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);

            foreach (var columnProfile in pageRangeColumnProfiles)
            {
                var numberOfColumns = columnProfile.Columns.Count;

                for (var page = columnProfile.StartPage; page <= columnProfile.EndPage; page++)
                {
                    if (Program.ExecutionProfile == ExecutionProfile.Testing)
                    {
                        if (page > 5)
                        {
                            Console.WriteLine("Finished testing. Exiting for loop.");
                            break;
                        }
                    }

                    var textFromPage = new string[numberOfColumns][];

                    for (var columnIndex = 0; columnIndex < numberOfColumns; columnIndex++)
                    {
                        var currentColumn = columnProfile.Columns[columnIndex];
                        textFromPage[columnIndex] = ExtractCurrentColumnFromPage(currentColumn, reader, page);
                    }

                    var textAsCsv = ConvertTextFromPageToCsv(textFromPage);

                    SaveCurrentPageToFile(textAsCsv, fileStream);

                    Console.WriteLine($"Converted page: {page}");
                }

                if (Program.ExecutionProfile == ExecutionProfile.Testing)
                {
                    Console.WriteLine("Exiting foreach loop.");
                    break;
                }
            }
        }

        private string[] ExtractCurrentColumnFromPage(Rectangle column, PdfReader reader, int page)
        {
            var renderFilter = new RegionTextRenderFilter(column);
            var renderFilterArray = new RenderFilter[] {renderFilter};

            var filteredTextRenderListener = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), renderFilterArray);
            var textFromColumn = PdfTextExtractor.GetTextFromPage(reader, page, filteredTextRenderListener);

            return SplitColumnTextIntoRows(textFromColumn);
        }

        private string[] SplitColumnTextIntoRows(string textFromColumn) => textFromColumn.Split(new[] {Environment.NewLine}, StringSplitOptions.None);

        private string ConvertTextFromPageToCsv(string[][] textFromPage)
        {
            var stringBuilder = new StringBuilder();

            int columnCount = textFromPage.Length, rowCount = textFromPage[0].Length;

            for (var i = 0; i < rowCount; i++)
            {
                for (var j = 0; j < columnCount; j++)
                {
                    var cell = textFromPage[j][i];
                    var text = CleanTextForCsv(cell, j);
                    stringBuilder.Append(text);
                    if (j < columnCount - 1)
                    {
                        stringBuilder.Append(", ");
                    }
                    else
                    {
                        stringBuilder.AppendLine();
                    }
                }
            }

            return stringBuilder.ToString();
        }

        private string CleanTextForCsv(string s, int columnIndex)
        {
            if (columnIndex == 2)
            {
                s = s.Replace(",", "");
            }

            if (!s.Contains(",")) return s;
            Console.WriteLine($"Found a comma here: {s}");
            return $@"""{s}""";
        }

        private void SaveCurrentPageToFile(string textAsCsv, FileStream fileStream)
        {
            var byteArray = Encoding.UTF8.GetBytes(textAsCsv);

            fileStream.Write(byteArray, 0, byteArray.Length);

            fileStream.Flush();
        }
    }
}