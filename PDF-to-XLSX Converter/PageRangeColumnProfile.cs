using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using iTextSharp.text;

namespace PDF_to_XLSX_Converter
{
    public class PageRangeColumnProfile
    {
        public PageRangeColumnProfile(int startPage, int endPage, List<float> columnEndsInInches)
        {
            StartPage = startPage;
            EndPage = endPage;
            Columns = new Collection<Rectangle>();

            CreateColumnRectangles(columnEndsInInches);

            Console.WriteLine($"Column profile created for pages {startPage} through {endPage}.");
        }

        public int StartPage { get; }
        public int EndPage { get; }
        public Collection<Rectangle> Columns { get; }

        private void CreateColumnRectangles(List<float> columnEndsInInches)
        {
            float lowerLeftY = InchesToPoints(0), upperRightY = InchesToPoints(8), lowerLeftX = InchesToPoints(0), upperRightX;

            foreach (var currentColumnEnding in columnEndsInInches.Select(InchesToPoints))
            {
                upperRightX = currentColumnEnding;
                Columns.Add(new Rectangle(lowerLeftX, lowerLeftY, upperRightX, upperRightY));;
                lowerLeftX = currentColumnEnding;
            }
            
            static float InchesToPoints(float inches) => inches * 72;
        }
    }
}