using System;
using System.Collections.Generic;
using static PDF_to_XLSX_Converter.ExecutionProfile;

namespace PDF_to_XLSX_Converter
{
    class Program
    {
        public static readonly ExecutionProfile ExecutionProfile = ConvertToXlsx;

        private static string _pdfFileLocation  = @"/Users/rjensen/Documents/ARKANSAS LIST UNCLAIMED.pdf";
        private static string _csvFileLocation  = @"/Users/rjensen/Documents/ARKANSAS LIST UNCLAIMED.csv";
        private static string _xlsxFileLocation = @"/Users/rjensen/Documents/ARKANSAS LIST UNCLAIMED.xlsx";

        private static List<PageRangeColumnProfile> _pageRangeColumnProfiles;

        public static void Main(string[] args)
        {
            CreatePageRangeProfiles();
            ConvertPdf();
        }

        private static void CreatePageRangeProfiles()
        {
            var columnEndings1  = new List<float> {2.26f, 5.62f, 6.01f, 8.62f, 9.05f, 10.41f, 11.00f};
            var columnEndings2  = new List<float> {2.28f, 5.75f, 6.21f, 8.75f, 9.14f, 10.53f, 11.00f};
            var columnEndings3  = new List<float> {2.19f, 5.88f, 6.31f, 8.95f, 9.25f, 10.54f, 11.00f};
            var columnEndings4  = new List<float> {2.33f, 5.67f, 6.11f, 8.71f, 9.04f, 10.49f, 11.00f};
            var columnEndings5  = new List<float> {2.31f, 5.86f, 6.27f, 8.81f, 9.16f, 10.53f, 11.00f};
            var columnEndings6  = new List<float> {2.19f, 5.78f, 6.22f, 8.80f, 9.18f, 10.60f, 11.00f};
            var columnEndings7  = new List<float> {2.30f, 5.76f, 6.20f, 8.78f, 9.16f, 10.58f, 11.00f};
            var columnEndings8  = new List<float> {2.31f, 5.79f, 6.23f, 8.82f, 9.19f, 10.59f, 11.00f};
            var columnEndings9  = new List<float> {2.23f, 5.79f, 6.21f, 8.79f, 9.12f, 10.52f, 11.00f};
            var columnEndings10 = new List<float> {2.75f, 7.56f, 8.08f, 8.49f, 8.79f, 10.43f, 11.00f};
            var columnEndings11 = new List<float> {2.90f, 7.45f, 8.00f, 8.36f, 8.73f, 10.52f, 11.00f};
            var columnEndings12 = new List<float> {2.27f, 5.79f, 6.28f, 8.68f, 9.07f, 10.44f, 11.00f};
            
            PageRangeColumnProfile profile1 = new PageRangeColumnProfile(1, 403, columnEndings1);
            PageRangeColumnProfile profile2 = new PageRangeColumnProfile(404, 994, columnEndings2);
            PageRangeColumnProfile profile3 = new PageRangeColumnProfile(995, 1500, columnEndings3);
            PageRangeColumnProfile profile4 = new PageRangeColumnProfile(1501, 2108, columnEndings4);
            PageRangeColumnProfile profile5 = new PageRangeColumnProfile(2109, 3305, columnEndings5);
            PageRangeColumnProfile profile6 = new PageRangeColumnProfile(3306, 3903, columnEndings6);
            PageRangeColumnProfile profile7 = new PageRangeColumnProfile(3904, 4498, columnEndings7);
            PageRangeColumnProfile profile8 = new PageRangeColumnProfile(4499, 5095, columnEndings8);
            PageRangeColumnProfile profile9 = new PageRangeColumnProfile(5096, 5693, columnEndings9);
            PageRangeColumnProfile profile10 = new PageRangeColumnProfile(5694, 6328, columnEndings10);
            PageRangeColumnProfile profile11 = new PageRangeColumnProfile(6329, 6671, columnEndings11);
            PageRangeColumnProfile profile12 = new PageRangeColumnProfile(6672, 7030, columnEndings12);

            _pageRangeColumnProfiles = new List<PageRangeColumnProfile>
            {
                profile1, profile2, profile3, profile4, profile5, profile6, 
                profile7, profile8, profile9, profile10, profile11, profile12
            };
        }

        private static void ConvertPdf()
        {
            switch (ExecutionProfile)
            {
                case ConvertToCsv:
                    new PdfToCsvConverter(_pdfFileLocation, _csvFileLocation)
                        .ExecuteOnColumnProfiles(_pageRangeColumnProfiles);
                    break;
                case ConvertToXlsx:
                    new PdfToXlsxConverter(_pdfFileLocation, _xlsxFileLocation)
                        .ExecuteOnColumnProfiles(_pageRangeColumnProfiles);
                    break;
            }
        }
    }
}