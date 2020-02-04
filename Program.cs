using System;
using System.IO;
using SautinSoft.Document;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Text;

namespace RTFToPDF
{
    class Program
    {
        static void Main(string[] args)
        {
            byte[] pdfBytes = File.ReadAllBytes(@"C:\Users\adit.shah\source\repos\aditshh\RTFToPDF\SampleFile\sampleWithImages.pdf");

            SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();

            f.OpenPdf(pdfBytes);

            if (f.PageCount > 0)
            {
                byte[] word = f.ToWord();
                System.IO.File.WriteAllBytes(@"C:\Users\adit.shah\source\repos\aditshh\RTFToPDF\SampleFile\sampleWithImages_sautinSoft.rtf", word);

                f.HtmlOptions.IncludeImageInHtml = true;
                f.HtmlOptions.InlineCSS = true;
                string html = f.ToHtml();

                System.IO.File.WriteAllText(@"C:\Users\adit.shah\source\repos\aditshh\RTFToPDF\SampleFile\sampleWithImages_sautinSoft.html", html);
                
                //Console.WriteLine(html);
                //Console.WriteLine("In base64 rtf: {0}", Convert.ToBase64String(word));

                string rez = GetText(@"C:\Users\adit.shah\source\repos\aditshh\RTFToPDF\SampleFile\sampleWithImages.pdf");
                System.IO.File.WriteAllText(@"C:\Users\adit.shah\source\repos\aditshh\RTFToPDF\SampleFile\sampleWithImages_iTextSharp.txt", rez);
                System.IO.File.WriteAllText(@"C:\Users\adit.shah\source\repos\aditshh\RTFToPDF\SampleFile\sampleWithImages_iTextSharp.rtf", rez);


                string rHTML = ConvertPdf(@"C:\Users\adit.shah\source\repos\aditshh\RTFToPDF\SampleFile\sampleWithImages.pdf", @"C:\Users\adit.shah\source\repos\aditshh\RTFToPDF\SampleFile\sampleWithImages_iTextSharp.html");


                //string spireReturn = SpirePDF(@"C:\Users\adit.shah\source\repos\aditshh\RTFToPDF\SampleFile\sampleWithImages.pdf", @"C:\Users\adit.shah\source\repos\aditshh\RTFToPDF\SampleFile\sampleWithImages_SpirePDF.html");
                Console.ReadLine();
            }
        }

        /*
        public static string SpirePDF(string inputFileName, string outputFileName)
        {
            //The Spire Returns a watermark
            PdfDocument document = new PdfDocument();
            document.LoadFromFile(inputFileName);

            //Save doc file to html
            document.SaveToFile(outputFileName, FileFormat.HTML);
            
            return "success";
        }
        */

        public static string ConvertPdf(string inputFileName, string outputFileName)
        {
            if (string.IsNullOrEmpty(inputFileName) || string.IsNullOrEmpty(outputFileName))
            {
                return "File name error";
            }
            else if (!File.Exists(inputFileName))
            {
                return "File is not exist";
            }
            else
            {
                PdfReader pr = new PdfReader(inputFileName);

                int maxPage = pr.NumberOfPages;
                pr.Close();
                StreamWriter outFile = new StreamWriter(outputFileName, true, System.Text.Encoding.UTF8);

                int page = 1;

                while (page <= maxPage)
                {
                    ITextExtractionStrategy its = new SimpleTextExtractionStrategy();
                    PdfReader reader = new PdfReader(inputFileName);
                    
                    outFile.Write(PdfTextExtractor.GetTextFromPage(reader, page, its));

                    reader.Close();
                    page++;

                }

                outFile.Close();
            }
            return "success";
        }


        /// <summary>
        /// Extracts a text from a PDF file.
        /// </summary>
        /// <param name="filePath">the full path to the pdf file.</param>
        /// <returns>the extracted text</returns>
        public static string GetText(string filePath)
        {
            var sb = new StringBuilder();
            try
            {
                using (PdfReader reader = new PdfReader(filePath))
                {
                    string prevPage = "";
                    for (int page = 1; page <= reader.NumberOfPages; page++)
                    {
                        ITextExtractionStrategy its = new SimpleTextExtractionStrategy();
                        var s = PdfTextExtractor.GetTextFromPage(reader, page, its);
                        if (prevPage != s) sb.Append(s);
                        prevPage = s;
                    }
                    reader.Close();
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            return sb.ToString();
        }
    }

}