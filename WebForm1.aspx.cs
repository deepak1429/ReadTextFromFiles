using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebApplication3
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            //string pdfText = GetTextFromPDF();

            Response.Write(GetTextFromPDF());
            //string wordText = GetTextFromWord();
            //string Text = GetTextFromText();  

        }

        /// <summary>
        /// Reading text from PDF
        /// </summary>
        /// <returns></returns>
        private string GetTextFromPDF()
        {
            StringBuilder text = new StringBuilder();
            using (PdfReader reader = new PdfReader(Server.MapPath("~/Files/Unbilled_Transactions.pdf")))
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                }
            }

            return text.ToString();
        }

        /// <summary>
        /// Reading Text from Word document
        /// </summary>
        /// <returns></returns>
        //private string GetTextFromWord()
        //{
        //    StringBuilder text = new StringBuilder();
        //    Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
        //    object miss = System.Reflection.Missing.Value;
        //    object path = @"D:\Articles2.docx";
        //    object readOnly = true;
        //    Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
        //    for (int i = 0; i < docs.Paragraphs.Count; i++)
        //    {
        //        text.Append(" \r\n " + docs.Paragraphs[i + 1].Range.Text.ToString());
        //    }

        //    return text.ToString();
        //}

        /// <summary>
        /// Reading text from text files
        /// </summary>
        /// <returns></returns>
        private string GetTextFromText()
        {
            string text = System.IO.File.ReadAllText(@"D:\Articles2.txt");

            return text.ToString();
        }
    }
}