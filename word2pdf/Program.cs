using System;
using System.IO;
using WordToPDF;

namespace word2pdf
{
    class Program
    {
        static void Main(string[] args)
        {
            string wordFile = args[0]; //@"C:\Temp\202103\17\1\BIOMAC_17904.docx";
            Word2PDF(wordFile);
        }

        public static void Word2PDF(string wordFile)
        {
            try
            {
                Word2Pdf objWorPdf = new Word2Pdf();
                string targetPath = Path.ChangeExtension(wordFile, ".pdf");
                if (File.Exists(targetPath)) File.Delete(targetPath);

                if (Path.GetExtension(wordFile) == ".doc" || Path.GetExtension(wordFile) == ".docx")
                {
                    object ToLocation = targetPath;
                    objWorPdf.InputLocation = wordFile;
                    objWorPdf.OutputLocation = ToLocation;
                    objWorPdf.Word2PdfCOnversion();
                }
            }
            catch (Exception ex)
            {

                throw;
            }
        }

    }
}
