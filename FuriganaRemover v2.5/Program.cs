using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Diagnostics;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using BitMiracle.Docotic;
using BitMiracle.Docotic.Pdf;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;


namespace FuriganaRemover_v2._5
{
    class Program
    {
        static void Main(string[] args)
        {
            // Changes \ to / in the path string
            string oldstr = @"\";
            string newstr = @"/";
            string pdf_path_original = string.Empty;

            // Tells the file path
            Console.WriteLine("FILE PATH: ");
            Console.OutputEncoding = Encoding.GetEncoding(932);
            Console.WriteLine("ファイル パス: ");
            pdf_path_original = Console.ReadLine();
            string pdf_path = pdf_path_original.Replace(oldstr, newstr);

            // Tells the file name
            Console.WriteLine(".pdf`s Name: ");
            Console.OutputEncoding = Encoding.GetEncoding(932);
            Console.WriteLine("ファイルの名前は: ");
            string pdf_file_name = Console.ReadLine();
            string path_n_name = pdf_path + "/" + pdf_file_name + ".pdf";

            // Tells the file alignment
            Console.WriteLine("Text alignment: ");
            Console.OutputEncoding = Encoding.GetEncoding(932);
            Console.WriteLine("テキスト配置: ");
            Console.WriteLine("1: Left (左), and Right (右), or 2: Middle (真ん中)");
            string caseSwitch = Console.ReadLine();

            Console.WriteLine();


            // Tells the file start page and end page
            Console.WriteLine(".pdf`s Initial Page ");
            Console.OutputEncoding = Encoding.GetEncoding(932);
            Console.WriteLine("最初のページは: ");
            string startpage_string = Console.ReadLine();
            int startpage = Int32.Parse(startpage_string);
            Console.WriteLine(".pdf`s Last Page ");
            Console.OutputEncoding = Encoding.GetEncoding(932);
            Console.WriteLine("最後のページは: ");
            string endpage_string = Console.ReadLine();
            int endpage = Int32.Parse(endpage_string);

            // Creates blank pdf files
            for (int i = startpage; i <= endpage; i++)
            {
                string convi = i.ToString();
                // Creates a temp folder for the pdf files
                System.IO.Directory.CreateDirectory(pdf_path + "/" + "temp");
                System.IO.FileStream pdf_creator = new System.IO.FileStream(pdf_path + "/" + "temp" + "/" + convi + ".pdf", System.IO.FileMode.Create);
                pdf_creator.Close();
            }


            // Tells the attributes from the new pdf files, and the original pdf source
            iTextSharp.text.pdf.PdfReader reader = null;
            iTextSharp.text.Document sourceDocument = null;
            iTextSharp.text.pdf.PdfCopy pdfCopyProvider = null;
            iTextSharp.text.pdf.PdfImportedPage importedPage = null;

            reader = new iTextSharp.text.pdf.PdfReader(path_n_name);
            sourceDocument = new iTextSharp.text.Document(reader.GetPageSizeWithRotation(startpage));
            sourceDocument.Open();

            // Creates a .docx to receive the pdf's text
            Spire.Doc.Document word_doc = new Spire.Doc.Document();

            // Word doc formatting
            Spire.Doc.Section word_doc_section = word_doc.AddSection();
            Spire.Doc.Documents.Paragraph word_doc_paragraph = word_doc_section.AddParagraph();
            Spire.Doc.Documents.Paragraph word_doc_paragraph_page = word_doc_section.AddParagraph();


            // Update those blank pdf files, inserting the copied pages into it
            try
            {

                for (int i = startpage; i <= endpage; i++)
                {
                    string convi = i.ToString();
                    pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(pdf_path + "/" + "temp" + "/" + convi + ".pdf", System.IO.FileMode.Append));
                    sourceDocument.Open();
                    importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                    pdfCopyProvider.AddPage(importedPage);

                }


                sourceDocument.Close();
                reader.Close();
            }


            // ERROR
            catch (Exception ex)
            {
                Console.WriteLine("Error! ");
                Console.OutputEncoding = Encoding.GetEncoding(932);
                Console.WriteLine("エラー ! ");
                throw ex;
            }


            // Collects the text without furigana from the listed pdf files
            switch (caseSwitch)
            {
                // case 1 reffers to the left and right alignments of the pdf text
                case "1":
                    Console.WriteLine();
                    for (int i = startpage; i <= endpage; i++)
                    {
                        // the following refers to the int counter of pages being converted into string
                        string convi = i.ToString();
                        Console.OutputEncoding = Encoding.GetEncoding(932);
                        Console.WriteLine("今のページ： " + convi);
                        Console.WriteLine("Current Page： " + convi);

                        // the following refers to the bitmiracle api pdf to get the texts
                        using (BitMiracle.Docotic.Pdf.PdfDocument pdf_1 = new BitMiracle.Docotic.Pdf.PdfDocument(pdf_path + "/" + "temp" + "/" + convi + ".pdf"))
                        {
                            BitMiracle.Docotic.Pdf.PdfPage page = pdf_1.Pages[0];
                            foreach (PdfTextData data in page.GetWords())
                            {

                                if (data.FontSize > 6 && data.Position.X < 600)
                                {
                                    string text = data.Text;
                                    text.TrimEnd();
                                    Console.OutputEncoding = Encoding.GetEncoding(932);
                                    Console.WriteLine(text);
                                    //word_builder.Writeln(text);
                                    word_doc_paragraph.AppendText(text);

                                }

                            }
                            foreach (PdfTextData data in page.GetWords())
                            {

                                if (data.FontSize > 6 && data.Position.X > 600)
                                {
                                    string text = data.Text;
                                    text.TrimEnd();
                                    Console.OutputEncoding = Encoding.GetEncoding(932);
                                    Console.WriteLine(text);
                                    word_doc_paragraph.AppendText(text);

                                }


                            }
                        }
                        // the following lines reffers to the space between pages of the pdf text
                        Console.WriteLine();
                        Console.WriteLine();
                        Console.WriteLine();
                        Console.WriteLine();
                        Console.WriteLine();
                        // the followin reffers to the extra lines on word text


                        word_doc_paragraph.AppendText("                                        ");
                        word_doc_paragraph.AppendText("CURRENT PAGE： " + convi);
                        word_doc_paragraph = word_doc_section.AddParagraph();
                        word_doc.Sections[0].Paragraphs[i].AppendBreak(BreakType.PageBreak);





                    }

                    break;


                // case 2 reffers to the alignment of the pdf text that is centralized
                case "2":
                    Console.WriteLine();
                    for (int i = startpage; i <= endpage; i++)
                    {
                        // the following refers to the int counter of pages being converted into string
                        string convi = i.ToString();
                        Console.OutputEncoding = Encoding.GetEncoding(932);
                        Console.WriteLine("今のページ： " + convi);
                        Console.WriteLine("Current Page： " + convi);

                        // the following refers to the bitmiracle api pdf to get the texts
                        using (BitMiracle.Docotic.Pdf.PdfDocument pdf_1 = new BitMiracle.Docotic.Pdf.PdfDocument(pdf_path + "/" + "temp" + "/" + convi + ".pdf"))
                        {
                            BitMiracle.Docotic.Pdf.PdfPage page = pdf_1.Pages[0];
                            foreach (PdfTextData data in page.GetWords())
                            {

                                if (data.FontSize > 6)
                                {
                                    string text = data.Text;
                                    text.TrimEnd();
                                    Console.OutputEncoding = Encoding.GetEncoding(932);
                                    Console.WriteLine(text);
                                    word_doc_paragraph.AppendText(text);


                                }


                            }

                        }
                        Console.WriteLine();
                        Console.WriteLine();
                        Console.WriteLine();
                        Console.WriteLine();
                        Console.WriteLine();

                        word_doc_paragraph.AppendText("                                        ");
                        word_doc_paragraph.AppendText("CURRENT PAGE： " + convi);
                        word_doc_paragraph = word_doc_section.AddParagraph();
                        word_doc.Sections[0].Paragraphs[i].AppendBreak(BreakType.PageBreak);



                    }
                    break;

                default:
                    Console.OutputEncoding = Encoding.GetEncoding(932);
                    Console.WriteLine("error! (エラー)");
                    Console.ReadKey();
                    break;
            }

            // The following refers to creating a .docx file, opening up the file and deleting the temp folder
            word_doc.SaveToFile(pdf_path + "/" + pdf_file_name + ".docx", FileFormat.Docx);
            System.IO.Directory.Delete(pdf_path + "/" + "temp", true);
            try
            {
                System.Diagnostics.Process.Start(pdf_path + "/" + pdf_file_name + ".docx");
            }
            catch
            {
                Console.WriteLine("Error! ");
                Console.OutputEncoding = Encoding.GetEncoding(932);
                Console.WriteLine("エラー ! ");

            }



        }



    }
}
