using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp2
{
    public class PDFFooter : PdfPageEventHelper
    {
        // write on top of document
        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            base.OnOpenDocument(writer, document);
            PdfPTable tabFot = new PdfPTable(new float[] { 1F });
            tabFot.SpacingAfter = 10F;
            PdfPCell cell;
            tabFot.TotalWidth = 300F;
            cell = new PdfPCell(new Phrase(""));
            cell.Border = Rectangle.NO_BORDER;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;    //置中

            tabFot.AddCell(cell);
            tabFot.WriteSelectedRows(0, -1, 150, document.Top, writer.DirectContent);
        }

        // write on start of each page
        public override void OnStartPage(PdfWriter writer, Document document)
        {
            base.OnStartPage(writer, document);
        }

        // write on end of each page
        public override void OnEndPage(PdfWriter writer, Document document)
        {
            //設定中文字體
            string chFontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "KAIU.TTF");
            BaseFont chBaseFont = BaseFont.CreateFont(chFontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font chFont = new iTextSharp.text.Font(chBaseFont, 14);

            // header設定
            Rectangle page = document.PageSize;
            PdfPTable head = new PdfPTable(1);
            head.TotalWidth = page.Width;
            Phrase complicatedPhrase = new Phrase();
            complicatedPhrase.Add(new Chunk("test \n", new iTextSharp.text.Font(chBaseFont, 18, 1)));
            complicatedPhrase.Add(new Chunk("TCI Co., Ltd.  ", new Font(Font.FontFamily.TIMES_ROMAN, 18, 1)));
            complicatedPhrase.Add(new Chunk("TPE/SH/JP/US/HK", new Font(Font.FontFamily.TIMES_ROMAN, 14, 1)));
            PdfPCell c = new PdfPCell(complicatedPhrase);
            c.Border = Rectangle.NO_BORDER;
            c.VerticalAlignment = Element.ALIGN_TOP;
            c.HorizontalAlignment = Element.ALIGN_CENTER;
            head.AddCell(c);
            head.WriteSelectedRows(
              // first/last row; -1 writes all rows
              0, -1,
              // left offset
              0,
              // ** bottom** yPos of the table
              page.Height - document.TopMargin + head.TotalHeight + 20,
              writer.DirectContent
            );

            //logo
            //將圖片加入到某個座標位置absolute position
            var imageFileName = @"test11\logo.jpg";
            iTextSharp.text.Image googleJPG1 = iTextSharp.text.Image.GetInstance(imageFileName);
            //調整圖片大小
            googleJPG1.ScalePercent(60f);
            googleJPG1.SetAbsolutePosition(65, 755);
            document.Add(googleJPG1);

            //畫圖
            PdfContentByte grx = writer.DirectContent;
            // add a rectangle
            //grx.Rectangle(100, 700, 100, 100);
            // add the diagonal
            grx.SetLineWidth(5);
            grx.SetColorStroke(BaseColor.LIGHT_GRAY);
            grx.MoveTo(40, 753);
            grx.LineTo(550, 753);
            // stroke the lines
            grx.Stroke();

            // footer設定
            PdfPTable footer = new PdfPTable(1);
            footer.TotalWidth = page.Width;
            Phrase phrase2 = new Phrase();
            phrase2.Add(new Chunk("test11 \n", new iTextSharp.text.Font(chBaseFont, 10)));
            phrase2.Add(new Chunk("test11 \n", new Font(Font.FontFamily.UNDEFINED, 10, 1)));
            phrase2.Add(new Chunk("test11 \n", new Font(Font.FontFamily.UNDEFINED, 10, 1)));

            Chunk underLineChunk = new Chunk("WWW.test11.COM ", new Font(Font.FontFamily.UNDEFINED, 10, 1));  //style =1:粗體
            underLineChunk.SetUnderline(0.2f, -2f);
            phrase2.Add(underLineChunk);
            PdfPCell c2 = new PdfPCell(phrase2);
            c2.Border = Rectangle.NO_BORDER;
            c2.VerticalAlignment = Element.ALIGN_TOP;
            c2.HorizontalAlignment = Element.ALIGN_CENTER;
            footer.AddCell(c2);
            footer.WriteSelectedRows(
              // first/last row; -1 writes all rows
              0, -1,
              // left offset
              0,
              // ** bottom** yPos of the table
              document.Bottom - 20,
              writer.DirectContent
            );
        }

        //write on close of document
        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            base.OnCloseDocument(writer, document);
        }
    }
}
