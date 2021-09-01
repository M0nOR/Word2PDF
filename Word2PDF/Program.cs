using System;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace Word2PDF
{
    class Program
    {
        static void Main(string[] files)
        {
            /*
            files = new string[] 
            {
                "D:/Test/Casos clinicos.docx",
                "D:/Test/Ejemplo Informe Pericial.docx"
            };*/
            Console.WriteLine(String.Join("\n", files));
            if (files.Length > 0)
            {
                Word._Application wordApp = new Word.Application();
                foreach (string file in files)
                {
                    FileInfo fileInfo = new FileInfo(file);
                    if (fileInfo.Exists)
                    {
                        try
                        {
                            Console.WriteLine("Abriendo archivo: {0}.", fileInfo.FullName);
                            wordApp.Documents.Open(fileInfo.FullName);
                            string namePDF = fileInfo.Name.Substring(0, fileInfo.Name.IndexOf(".")) + ".pdf";
                            string newName = fileInfo.DirectoryName + "/" + namePDF;

                            wordApp.ActiveDocument.SaveAs2(
                                newName,
                                FileFormat: Word.WdSaveFormat.wdFormatPDF);
                            Console.WriteLine("Archivo convertido: {0}.", newName);
                            wordApp.ActiveDocument.Close(SaveChanges: false);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                            Console.ReadKey();
                        }
                    }
                }
                wordApp.Quit();

                Console.WriteLine("\n\n\nMade by: Mon Ortiz :)");
                //Console.SetBufferSize(400, 400);

                Console.WriteLine("Presione una tecla para continuar...");
                Console.ReadKey();
            }
        }
    }
}
