using System;
using System.IO;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas.Parser.Data;

namespace PdfToText
{
    class Program
    {
        static void Main(string[] args)
        {
            // Hardcoded input PDF file and output text file paths
            string inputPdfFile = @"C:\books\book.pdf";
            string outputTextFile = @"C:\books\book.txt";

            try
            {
                // Extract text from the input PDF file
                string extractedText = ExtractTextFromPdf(inputPdfFile);
                // Write the extracted text to the output text file
                File.WriteAllText(outputTextFile, extractedText);
                Console.WriteLine($"Text extracted and saved to: {outputTextFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error extracting text from PDF: {ex.Message}");
            }
        }

        // Extracts text from the input PDF file, filtering out unwanted content
        static string ExtractTextFromPdf(string pdfFile)
        {
            using PdfReader pdfReader = new PdfReader(pdfFile);
            using PdfDocument pdf = new PdfDocument(pdfReader);

            // Find the starting and ending pages of the main content
            int startPageIndex = FindStartPage(pdf);
            int endPageIndex = FindEndPage(pdf);

            if (startPageIndex < 0 || endPageIndex < 0)
            {
                throw new InvalidOperationException("Could not find the Introduction and/or the last chapter/epilogue.");
            }

            var strategy = new FilteredTextExtractionStrategy();
            string result = "";

            // Iterate through the main content pages, extracting text within the specified margins
            for (int i = startPageIndex; i <= endPageIndex; i++)
            {
                strategy.SetExtractionArea(pdf.GetPage(i).GetPageSize().ApplyMargins(LeftMargin, RightMargin, TopMargin, BottomMargin, false));
                string pageText = PdfTextExtractor.GetTextFromPage(pdf.GetPage(i), strategy);
                result += pageText + Environment.NewLine;
            }

            return result;
        }

        // Constants for the margins, adjust these as necessary
        private const float LeftMargin = 50;
        private const float RightMargin = 50;
        private const float TopMargin = 80;
        private const float BottomMargin = 80;

        // Finds the starting page of the main content by searching for the "Introduction" or a similar keyword
        static int FindStartPage(PdfDocument pdf)
        {
            for (int i = 1; i <= pdf.GetNumberOfPages(); i++)
            {
                var text = PdfTextExtractor.GetTextFromPage(pdf.GetPage(i));
                if (text.Contains("Introduction"))
                {
                    return i;
                }
            }

            return -1;
        }

        // Finds the ending page of the main content by searching for the "Epilogue", "Conclusion", or a similar keyword
        static int FindEndPage(PdfDocument pdf)
        {
            for (int i = pdf.GetNumberOfPages(); i >= 1; i--)
            {
                var text = PdfTextExtractor.GetTextFromPage(pdf.GetPage(i));
                if (text.Contains("Epilogue") || text.Contains("Conclusion"))
                {
                    return i;
                }
            }

            return -1;
        }
    }


    // Custom text extraction strategy that filters out text outside the specified extraction area
    class FilteredTextExtractionStrategy : LocationTextExtractionStrategy
    {
        private Rectangle extractionArea;
                    // Set the extraction area for filtering unwanted content
        public void SetExtractionArea(Rectangle extractionArea)
        {
            this.extractionArea = extractionArea;
        }

        // Override the EventOccurred method to filter out text outside the extraction area
        public override void EventOccurred(IEventData data, EventType type)
        {
            if (type.Equals(EventType.RENDER_TEXT))
            {
                TextRenderInfo renderInfo = (TextRenderInfo)data;
                foreach (TextRenderInfo ri in renderInfo.GetCharacterRenderInfos())
                {
                    Vector start = ri.GetBaseline().GetStartPoint();
                    Vector end = ri.GetBaseline().GetEndPoint();
                    float x = Math.Min(start.Get(Vector.I1), end.Get(Vector.I1));
                    float y = Math.Min(start.Get(Vector.I2), end.Get(Vector.I2));
                    float width = Math.Abs(end.Get(Vector.I1) - start.Get(Vector.I1));
                    float height = Math.Abs(end.Get(Vector.I2) - start.Get(Vector.I2));

                    Rectangle textRectangle = new Rectangle(x, y, width, height);

                    // Include only text within the extraction area
                    if (extractionArea.Contains(textRectangle))
                    {
                        base.EventOccurred(ri, type);
                    }
                }
            }
            else
            {
                base.EventOccurred(data, type);
            }
        }
    }
}
