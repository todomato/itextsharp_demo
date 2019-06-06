using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace ConsoleApp2
{
    //製作PDF by itextsharp
    class Program
    {
        static void Main(string[] args)
        {
            //設定文件紙張大小
            Document doc = new Document(PageSize.A4, 27, 27, 100, 100);

            //設定pdf路徑
            string path = @"C:\Users\xxxxx\bin\Debug\";
            if (Directory.Exists(path) == false)
            {
                Directory.CreateDirectory(path);
            }

            //設定pdf檔案名稱
            MemoryStream outputStream = new MemoryStream();//要把PDF寫到哪個串流

            var guid = Guid.NewGuid();
            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(path + guid + ".pdf", FileMode.Create));
            writer.PageEvent = new PDFFooter();
            doc.Open();

            //Raw Material Specification
            Chunk underLineChunk = new Chunk("xxxx", new Font(Font.FontFamily.UNDEFINED, 12, 1));  //style =1:粗體
            underLineChunk.SetUnderline(0.2f, -2f);
            Phrase myPhrase = new Phrase();
            myPhrase.Add(underLineChunk);
            Paragraph myPara = new Paragraph();
            myPara.Alignment = Element.ALIGN_CENTER;
            myPara.Add(myPhrase);
            doc.Add(myPara);

            //空白
            var emptyParagraph = new Paragraph();
            emptyParagraph.Add("\n");
            doc.Add(emptyParagraph);

            //Issue Date:
            var issueDate = DateTime.Now.ToString("MMM.dd.yyyy", new CultureInfo("en-US"));
            myPhrase = new Phrase();
            myPhrase.Add(new Chunk("Issue Date:" + issueDate, new Font(Font.FontFamily.UNDEFINED, 11)));
            myPara = new Paragraph();
            myPara.Alignment = Element.ALIGN_RIGHT;
            myPara.SpacingAfter = 3;
            myPara.Add(myPhrase);
            doc.Add(myPara);

            // 表格
            PdfPTable table = new PdfPTable(new float[] { 1f, 1.4f, 1f, 1.1f });
            table.WidthPercentage = 100f;
            var font = new Font(Font.FontFamily.UNDEFINED, 11);
            //
            table.AddCell(new myPdfPCell(new Phrase(" Commodity", font)));
            table.AddCell(new myPdfPCell(new Phrase("")));
            table.AddCell(new myPdfPCell(new Phrase(" Raw Material Code", font)));
            table.AddCell(new myPdfPCell(new Phrase("")));

            //
            table.AddCell(new myPdfPCell(new Phrase(" Common Name", font)));
            table.AddCell(new myPdfPCell(new Phrase("")));
            table.AddCell(new myPdfPCell(new Phrase(" Country of Origin", font)));
            table.AddCell(new myPdfPCell(new Phrase("")));

            //
            table.AddCell(new myPdfPCell(new Phrase(" Scientific Name", font)));
            table.AddCell(new myPdfPCell(new Phrase("")));
            table.AddCell(new myPdfPCell(new Phrase(" Shelf Life", font)));
            table.AddCell(new myPdfPCell(new Phrase("")));

            //
            myPdfPCell cell1 = new myPdfPCell(new Paragraph("Test Item", new Font(Font.FontFamily.UNDEFINED, 11, 1)));
            cell1.Colspan = 2;
            cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            table.AddCell(cell1);
            cell1 = new myPdfPCell(new Paragraph("Specification", new Font(Font.FontFamily.UNDEFINED, 11, 1)));
            cell1.Colspan = 2;
            cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            table.AddCell(cell1);

            //Physical Properties
            myPdfPCell rows = new myPdfPCell(new Phrase(" Physical Properties", font));
            rows.VerticalAlignment = Element.ALIGN_MIDDLE;
            rows.Rowspan = 5;
            table.AddCell(rows);
            //Appearance 
            table.AddCell(new myPdfPCell(new Phrase(" Appearance", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            //Appearance 
            table.AddCell(new myPdfPCell(new Phrase(" Brix", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            //Appearance 
            table.AddCell(new myPdfPCell(new Phrase(" pH", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            //Appearance 
            table.AddCell(new myPdfPCell(new Phrase(" Acidity", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            //Appearance 
            table.AddCell(new myPdfPCell(new Phrase(" Solubility", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            table.CompleteRow();

            //Analysis of Microorganisms
            rows = new myPdfPCell(new Phrase(" Analysis of \n Microorganisms", font));
            rows.VerticalAlignment = Element.ALIGN_MIDDLE;
            rows.Rowspan = 4;
            table.AddCell(rows);
            //Total plate count 
            table.AddCell(new myPdfPCell(new Phrase(" Total plate count", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            //Coliform 
            table.AddCell(new myPdfPCell(new Phrase(" Coliform", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            //E coli 
            table.AddCell(new myPdfPCell(new Phrase(" E. Coli", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            //Mold and Yeast 
            table.AddCell(new myPdfPCell(new Phrase(" Mold and Yeast", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            table.CompleteRow();

            //Heavy Metals
            rows = new myPdfPCell(new Phrase(" Heavy Metals", font));
            rows.VerticalAlignment = Element.ALIGN_MIDDLE;
            rows.Rowspan = 9;
            table.AddCell(rows);
            //as
            table.AddCell(new myPdfPCell(new Phrase(" As", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            //cd
            table.AddCell(new myPdfPCell(new Phrase(" Cd", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            //cr
            table.AddCell(new myPdfPCell(new Phrase(" Cr", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            //cu
            table.AddCell(new myPdfPCell(new Phrase(" Cu", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);

            //
            table.AddCell(new myPdfPCell(new Phrase(" Ge", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);

            //
            table.AddCell(new myPdfPCell(new Phrase(" Hg", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            //
            table.AddCell(new myPdfPCell(new Phrase(" Pb", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            //
            table.AddCell(new myPdfPCell(new Phrase(" Sb", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            //
            table.AddCell(new myPdfPCell(new Phrase(" Sn", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            table.CompleteRow();

            //Preservatives
            rows = new myPdfPCell(new Phrase(" Preservatives", font));
            rows.VerticalAlignment = Element.ALIGN_MIDDLE;
            rows.Rowspan = 3;
            table.AddCell(rows);
            //as
            table.AddCell(new Phrase(" Sorbic Acid (SA)", font));
            cell1 = new myPdfPCell(new Paragraph("N.D."));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            //cd
            table.AddCell(new Phrase(" Dehydroacetic Acid (DHA)", font));
            cell1 = new myPdfPCell(new Paragraph("N.D."));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            //cd
            table.AddCell(new Phrase(" Sodium Benzoate", font));
            cell1 = new myPdfPCell(new Paragraph("N.D."));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            table.CompleteRow();

            //Pesticide Residue
            rows = new myPdfPCell(new Phrase(" Pesticide Residue", font));
            rows.VerticalAlignment = Element.ALIGN_MIDDLE;
            rows.Rowspan = 2;
            table.AddCell(rows);
            //Carbamate
            table.AddCell(new myPdfPCell(new Phrase(" Carbamate", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            //Organophosphate
            table.AddCell(new myPdfPCell(new Phrase(" Organophosphate", font)));
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 2;
            table.AddCell(cell1);
            table.CompleteRow();

            //Pesticide Residue
            rows = new myPdfPCell(new Phrase(" Storage", font));
            rows.VerticalAlignment = Element.ALIGN_MIDDLE;
            table.AddCell(rows);
            //Carbamate
            cell1 = new myPdfPCell(new Paragraph(""));
            cell1.Colspan = 3;
            table.AddCell(cell1);
            table.CompleteRow();

            //將phrase內容加入到pdf檔案
            Paragraph para = new Paragraph();
            para.Add(table);
            para.Add("\n");

            doc.Add(para);
            doc.Close();

            var rpath =  guid + ".pdf";
            AddWaterMark(path, rpath);
        }

        private static void AddWaterMark(string prefix, string rpath)
        {
            
                using (PdfReader pdfReader = new PdfReader(prefix + rpath))
                {
                    int numberOfPages = pdfReader.NumberOfPages;
                    FileStream outputStream = new FileStream(prefix + "QQ" + rpath, FileMode.Create);

                    using (PdfStamper stamp = new PdfStamper(pdfReader, outputStream))
                    {
                        BaseFont bfChinese = BaseFont.CreateFont(@"C:\WINDOWS\Fonts\kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                        PdfGState gstate = new PdfGState()
                        {
                            FillOpacity = 0.25f,
                            StrokeOpacity = 0.25f
                        };

                        for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                        {
                            PdfContentByte pdfPageContents = stamp.GetOverContent(page);
                            pdfPageContents.SetGState(gstate); //塞入我們設定的透明度
                            var imageUrl = @"C:\Users\andy.chen\source\repos\ConsoleApp2\ConsoleApp2\Content\logo.png";
                            iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(imageUrl);
                            img.ScalePercent(100f);  //縮放比例
                            img.RotationDegrees = 40; //旋轉角度
                            img.SetAbsolutePosition(10, 240); //設定圖片每頁的絕對位置
                            PdfContentByte waterMark = stamp.GetOverContent(page);
                            waterMark.AddImage(img); //把圖片印上去 

                            // 加入所需文字和一些設定
                            //float fontSize = 100;
                            //float adjustTxtLineHeight = fontSize - 4;
                            //pdfPageContents.BeginText();
                            //pdfPageContents.SetFontAndSize(bfChinese, fontSize);
                            //pdfPageContents.SetRGBColorFill(210, 210, 210);
                            //pdfPageContents.SetGState(gstate);
                            //pdfPageContents.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TCI Test", 220, 400, 45);
                        }
                    }
                }
        }
    }

    public class myPdfPCell : PdfPCell
    {
        public myPdfPCell(Phrase phrase) : base(phrase)
        {
            this.PaddingBottom = 5;
        }
    }
}
