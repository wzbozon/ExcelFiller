using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using System.Web;
using System.IO;
using System.Net;
using System.Collections;
using System.Collections.Specialized;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace CSharpExcelWeb
{
    public partial class Form1 : Form
    {

        string[] authorsArray = new string[1000];
        string content;
        int countAuthors = 0;
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range range;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string url = "";
            string authorName = "";
            int cellRow;

            OpenExcelWorksheet();

            //Считываем массив авторов из Excel
            ReadFromExcel(2, 845);

            // Очищаем все строки от 2 до 2000
            ClearExcel(845);

            //Обновляем авторов, чтобы не повторялись
            for (int i = 1; i <= this.countAuthors; i++)
            {
                try
                {
                    UpdateExcelAuthors(i);
                }
                catch (Exception theException)
                {
                    String errorMessage;
                    errorMessage = "Error: ";
                    errorMessage = String.Concat(errorMessage, theException.Message);
                    errorMessage = String.Concat(errorMessage, " Line: ");
                    errorMessage = String.Concat(errorMessage, theException.Source);
                    MessageBox.Show(errorMessage, "Error");
                    continue;
                }
            }
            //MessageBox.Show(this.authorsArray[0], "Element #1");

            //Составляем запрос для википедии для первого автора для начала
            //Например, http://en.wikipedia.org/wiki/Edward_John_Phelps
            string wikiurl = "http://en.wikipedia.org/wiki/";
            //url = wikiurl + "Edward_John_Phelps";
            for (int i = 0; i <= this.countAuthors - 1; i++)
            {
                try
                {
                    authorName = this.authorsArray[i];
                    authorName = authorName.Trim();
                    authorName = authorName.Replace(" ", "_");
                    url = wikiurl + authorName;
                    //url = wikiurl + "Edward_John_Phelps";

                    ExtractContent(url);
                    CleanFromTags();
                    /*MessageBox.Show(this.authorsArray[i]);
                    MessageBox.Show(this.content);*/
                    cellRow = i + 2;
                    InsertIntoExcel(cellRow);
                }
                catch (Exception theException)
                {
                    /*String errorMessage;
                    errorMessage = "Error: ";
                    errorMessage = String.Concat(errorMessage, theException.Message);
                    errorMessage = String.Concat(errorMessage, " Line: ");
                    errorMessage = String.Concat(errorMessage, theException.Source);
                    MessageBox.Show(errorMessage, "Error");*/
                    continue;
                }
            }
        }

        public void ExtractContent(string url)
        {
            string text;
            string strURL = url;
            text = GetHtmlFromUrl(strURL);
            int begin;
            int end;
            int length;
            begin = text.IndexOf("<p>");
            end = text.IndexOf("</p>");
            length = end - begin + 4;
            text = text.Substring(begin, length);

            this.content = text;
        }

        private void CleanFromTags()
        {
            this.content = HtmlRemoval.StripTagsCharArray(this.content);
        }

        public enum ResponseCategories
        {
            Unknown = 0,       // Unknown code ( < 100 or > 599)
            Informational = 1, // Informational codes (100 >= 199)
            Success = 2,       // Success codes (200 >= 299)
            Redirected = 3,    // Redirection code (300 >= 399)
            ClientError = 4,   // Client error code (400 >= 499)
            ServerError = 5    // Server error code (500 >= 599)
        }

        public static string GetHtmlFromUrl(string url)
        {
            if (string.IsNullOrEmpty(url))
                throw new ArgumentNullException("url", "Parameter is null or empty");

            string html = "";
            HttpWebRequest request = GenerateHttpWebRequest(url);
            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                if (VerifyResponse(response) == ResponseCategories.Success)
                {
                    // Get the response stream.
                    Stream responseStream = response.GetResponseStream();
                    // Use a stream reader that understands UTF8.
                    using (StreamReader reader =
                    new StreamReader(responseStream, Encoding.UTF8))
                    {
                        html = reader.ReadToEnd();
                    }
                }
            }
            return html;
        }

        public static HttpWebRequest GenerateHttpWebRequest(string UriString)
        {
            // Get a Uri object.
            Uri Uri = new Uri(UriString);
            // Create the initial request.
            HttpWebRequest httpRequest = (HttpWebRequest)WebRequest.Create(Uri);
            // Return the request.
            return httpRequest;
        }

        // POST overload
        public static HttpWebRequest GenerateHttpWebRequest(string UriString,
            string postData,
            string contentType)
        {
            // Get a Uri object. 
            Uri Uri = new Uri(UriString);
            // Create the initial request.
            HttpWebRequest httpRequest = (HttpWebRequest)WebRequest.Create(Uri);

            // Get the bytes for the request; should be pre-escaped.
            byte[] bytes = Encoding.UTF8.GetBytes(postData);

            // Set the content type of the data being posted.
            httpRequest.ContentType = contentType;
            //"application/x-www-form-urlencoded"; for forms

            // Set the content length of the string being posted.
            httpRequest.ContentLength = postData.Length;

            // Get the request stream and write the post data in.
            using (Stream requestStream = httpRequest.GetRequestStream())
            {
                requestStream.Write(bytes, 0, bytes.Length);
            }
            // Return the request.
            return httpRequest;
        }

        public static ResponseCategories VerifyResponse(HttpWebResponse httpResponse)
        {
            // Just in case there are more success codes defined in the future
            // by HttpStatusCode, we will check here for the "success" ranges
            // instead of using the HttpStatusCode enum as it overloads some
            // values.
            int statusCode = (int)httpResponse.StatusCode;
            if ((statusCode >= 100) && (statusCode <= 199))
            {
                return ResponseCategories.Informational;
            }
            else if ((statusCode >= 200) && (statusCode <= 299))
            {
                return ResponseCategories.Success;
            }
            else if ((statusCode >= 300) && (statusCode <= 399))
            {
                return ResponseCategories.Redirected;
            }
            else if ((statusCode >= 400) && (statusCode <= 499))
            {
                return ResponseCategories.ClientError;
            }
            else if ((statusCode >= 500) && (statusCode <= 599))
            {
                return ResponseCategories.ServerError;
            }
            return ResponseCategories.Unknown;
        }

        private void OpenExcelWorksheet()
        {

            object misValue = System.Reflection.Missing.Value;

            // Specify a "currently active folder"
            string activeDir = @"c:\";

            string newPath = activeDir;

            // Create a new file name.

            string newFileName = "BusinessAuthors.xlsx";

            // Combine the new file name with the path
            newPath = System.IO.Path.Combine(newPath, newFileName);

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

        }

        private string[] ReadFromExcel(int fromNum, int toNum)
        {
            string fromRange, toRange;
            string str;
            string lastAuthor = "smth";
            int rCnt = 0;
            int cCnt = 0;
            this.countAuthors = 0;
            fromRange = "B" + fromNum.ToString();
            toRange = "B" + toNum.ToString();

            try
            {
                xlApp.Visible = true;

                range = xlWorkSheet.get_Range(fromRange, toRange);

                for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                {
                    for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                    {
                        str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                        if (str == lastAuthor) continue;
                        //MessageBox.Show(str);
                        this.countAuthors = this.countAuthors + 1;
                        this.authorsArray[countAuthors - 1] = str;
                        lastAuthor = str;
                    }
                }

                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                xlApp.Visible = true;
                xlApp.UserControl = true;

                string countAuthorsStr = countAuthors.ToString();
                MessageBox.Show(countAuthorsStr);
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
            return authorsArray;
        }

        private void InsertIntoExcel(int cellRow)
        {
            string fromRange = "C" + cellRow.ToString();
            string toRange = "C" + cellRow.ToString();

            try
            {
                xlApp.Visible = true;
                range = xlWorkSheet.get_Range(fromRange, toRange);
                range.Value2 = this.content;

                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
                MessageBox.Show(errorMessage, "Error");
                return;
            }
        }

        private void UpdateExcelAuthors(int authorNumber)
        {
            int cellRow = authorNumber + 1;
            string fromRange = "B" + cellRow.ToString();
            string toRange = "B" + cellRow.ToString();

            try
            {
                xlApp.Visible = true;
                range = xlWorkSheet.get_Range(fromRange, toRange);
                range.Value2 = this.authorsArray[authorNumber - 1];

                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
                MessageBox.Show(errorMessage, "Error");
                return;
            }
        }

        private void ClearExcel(int cellRow)
        {
            string fromRange = "B2";
            string toRange = "B" + cellRow.ToString();

            try
            {
                xlApp.Visible = true;
                range = xlWorkSheet.get_Range(fromRange, toRange);
                range.Value2 = "";

                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception theException)
            {
                /*String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
                MessageBox.Show(errorMessage, "Error");*/
                return;
            }
        }
    }
}
