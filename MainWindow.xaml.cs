using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Net;
using System.Net.Mail;


namespace DLPsetup
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public string filePathInput;
        public string clientName;
        //public List<string> keyTerms_titles = new List<string>();
        public List<string> keyTerms_contents = new List<string>();
        public string keyTerms_properties;
        public string combinedTitleNames;

        public MainWindow()
        {
            InitializeComponent();
            
        }

        private void MSWordButton_Click(object sender, EventArgs e)
        {
            CheckPathInput();
            CheckClientName();
            if (CheckKeyTermsEntered_FileNames())
            {
                //CreateDocument();
                CreateDocument_FileNames();
                CreateExcel_FileNames();
                CreatePPT_FileNames();
            }
            if (CheckKeyTermsEntered_FileContents())
            {
                CreateDocuments_contents();
                CreateExcel_contents();
                CreatePPT_contents();
                CreatePDF_all();
            }
            MessageBox.Show("All Documents should be created successfully!");
        }

        private void CheckPathInput()
        {
            filePathInput = FilePathTextBox.Text;
            if (string.IsNullOrEmpty(filePathInput))
            {
                MessageBox.Show("Enter Path Name where the documents should be stored!");
            }
            //else
            //{
            //    clientName = ClientNameTextBox.Text;
            //}
        }
        private void CheckClientName()
        {
            clientName = ClientNameTextBox.Text;
            if (string.IsNullOrEmpty(clientName))
            {
                MessageBox.Show("Enter Client Name");
            }
            //else
            //{
            //    clientName = ClientNameTextBox.Text;
            //}
        }
        private bool CheckKeyTermsEntered_FileNames()
        {
            combinedTitleNames = "";
            bool result = false;
            string input = FileNamesTextBox.Text;
            if (string.IsNullOrEmpty(input))
            {
                MessageBox.Show("Key Terms need to be entered!");
            }
            else
            {
                string lineStr = "";
                StringCollection lines = new StringCollection();
                int lineCount = FileNamesTextBox.LineCount;

                for (int line = 0; line < lineCount; line++)
                {
                    lines.Add(FileNamesTextBox.GetLineText(line));
                    //lineStr = SearchTermTextBox.GetLineText(line);
                    lineStr = FileNamesTextBox.GetLineText(line).Trim();
                    if (!string.IsNullOrEmpty(lineStr))
                    {
                        combinedTitleNames += lineStr + " ";
                        //keyTerms_titles.Add(FileNamesTextBox.GetLineText(line));
                    }
                }
                result = true;
            }
            return result;
        }
        private bool CheckKeyTermsEntered_FileContents()
        {
            keyTerms_properties = "";
            bool result = false;
            string input = DocumentContentsTextBox.Text;
            if (string.IsNullOrEmpty(input))
            {
                MessageBox.Show("Key Terms need to be entered!");
            }
            else
            {
                int lineCount = DocumentContentsTextBox.LineCount;

                for (int line = 0; line < lineCount; line++)
                {
                    keyTerms_contents.Add(DocumentContentsTextBox.GetLineText(line));
                    if (!(string.IsNullOrEmpty(DocumentContentsTextBox.GetLineText(line).Trim())))
                    {
                        keyTerms_properties += DocumentContentsTextBox.GetLineText(line).Trim() + " ";
                    }
                }
                result = true;
            }
            return result;
        }

        // Create document method 
        //private void CreateDocument()
        //{
        //    try
        //    {
        //        // Create instance for word app 
        //        Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();

        //        // Set animation status for word application
        //        winword.ShowAnimation = false;

        //        // Set status for word application visible or not 
        //        winword.Visible = false;

        //        // Create missing variable for missing value 
        //        object missing = System.Reflection.Missing.Value;

        //        // Create new document 
        //        Microsoft.Office.Interop.Word.Document document = winword.Documents.Add();

        //        // Add header into the document
        //        foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
        //        {
        //            // Get the header range and add the header details
        //            Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
        //            headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
        //            headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
        //            headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
        //            headerRange.Font.Size = 10;
        //            headerRange.Text = "Header text goes here";
        //        }

        //        // Add footers into the document 
        //        foreach (Microsoft.Office.Interop.Word.Section wordSection in document.Sections)
        //        {
        //            // Get the footer range and add the footer details 
        //            Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
        //            footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
        //            footerRange.Font.Size = 10;
        //            footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
        //            footerRange.Text = "Footer text goes here";
        //        }

        //        // Adding text to document
        //        document.Content.SetRange(0, 0);
        //        document.Content.Text = "This is a test document" + Environment.NewLine;

        //        // Add paragraph with Heading 1 style
        //        Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
        //        object styleHeading1 = "Heading 1";
        //        para1.Range.set_Style(ref styleHeading1);
        //        para1.Range.Text = "Para 1 text";
        //        para1.Range.InsertParagraphAfter();

        //        // Add paragraph with Heading 2 style
        //        Microsoft.Office.Interop.Word.Paragraph para2 = document.Content.Paragraphs.Add(ref missing);
        //        object styleHeading2 = "Heading 2";
        //        para2.Range.set_Style(ref styleHeading2);
        //        para2.Range.Text = "Para 2 text";
        //        para2.Range.InsertParagraphAfter();

        //        // Save document
        //        SaveDocument(document);

        //        winword.Quit(ref missing, ref missing, ref missing);
        //        winword = null;
        //        document.Close();

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}
        //private void SaveDocument(Microsoft.Office.Interop.Word.Document document)
        //{
        //    //object filename2 = @"C:\Avanade\temp1.docx";
        //    //Console.WriteLine(filename2.GetType());
        //    //document.SaveAs2(ref filename2);
        //    //document.Close();
        //    //document = null;
        //    ////MessageBox.Show("Document created successfully you beautiful person!");

        //    object filename;
        //    string keyTermStr;
        //    foreach(string keyTerm in keyTerms_titles)
        //    {
        //        //keyTerm.Replace("\r", string.Empty).Replace("\n", string.Empty);
        //        keyTermStr = keyTerm.Trim();
        //        filename = @"C:\Avanade\" + keyTermStr + ".docx";
        //        document.SaveAs2(ref filename);
        //        //document.SaveAs2(@"C:\Avanade\" + keyTermStr, ".docx");
        //        //document.Close();
        //        document = null;
        //        filename = null;
        //        keyTermStr = "";
        //    }
        //    MessageBox.Show("All Documents should be created successfully!");
        //}

        private void CreateDocument_FileNames()
        {
            try
            {
                // Create missing variable for missing value 
                object missing = System.Reflection.Missing.Value;

                // Create instance for word app 
                Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();

                // Set animation status for word application
                winword.ShowAnimation = false;

                // Set status for word application visible or not 
                winword.Visible = false;
                object filename;

                // SECTION BELOW CREATES 1 DOCUMENT CONTAINING ALL FILE NAMES 
                Microsoft.Office.Interop.Word.Document document = winword.Documents.Add();
                var newFilePath = Directory.CreateDirectory(filePathInput + "\\Word\\");
                filename = newFilePath.FullName + combinedTitleNames + ".docx";
                document.SaveAs2(ref filename);
                filename = null;

                document.Close();

                // SECTION BELOW CREATES NEW DOCUMENT FOR EACH FILE NAME 
                //string keyTermStr;
                //foreach (string keyTerm in keyTerms_titles)
                //{

                //    // Create new document 
                //    Microsoft.Office.Interop.Word.Document document = winword.Documents.Add();

                //    // Data
                //    keyTermStr = keyTerm.Trim();
                //    //filename = @"C:\Avanade\" + keyTermStr + ".docx";
                //    var newFilePath = Directory.CreateDirectory(filePathInput + "\\Word\\");
                //    filename = newFilePath.FullName + keyTermStr + ".docx";
                //    document.SaveAs2(ref filename);
                //    filename = null;
                //    keyTermStr = "";

                //    document.Close();
                //}
                //MessageBox.Show("All Documents should be created successfully!");
                winword.Quit(ref missing, ref missing, ref missing);
                //winword = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreateDocuments_contents()
        {
            CreateDocuments_contents_header();
            CreateDocuments_contents_body();
            CreateDocuments_contents_footer();
            CreateDocument_properties_category();
            CreateDocument_properties_keywords();
        }
        private void CreateDocuments_contents_header()
        {
            try
            {
                string headerText = "";
                // Create instance for word app 
                Microsoft.Office.Interop.Word.Application winword_header = new Microsoft.Office.Interop.Word.Application();

                // Set animation status for word application
                winword_header.ShowAnimation = false;

                // Set status for word application visible or not 
                winword_header.Visible = false;

                // Create missing variable for missing value 
                object missing = System.Reflection.Missing.Value;

                // Create new document 
                Microsoft.Office.Interop.Word.Document document_header = winword_header.Documents.Add();

                // Add header into the document
                foreach (Microsoft.Office.Interop.Word.Section section in document_header.Sections)
                {
                    // Get the header range and add the header details
                    Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                    headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
                    headerRange.Font.Size = 10;
                    foreach(string line in keyTerms_contents)
                    {
                        headerText += line;
                    }
                    headerRange.Text = keyTerms_properties;

                }

                // Save document
                //object filename = @"C:\Avanade\" + clientName + "_header.docx";
                var newFilePath = Directory.CreateDirectory(filePathInput + "\\Word\\");
                object filename = newFilePath.FullName + clientName + "_header.docx";
                document_header.SaveAs2(ref filename);
                //document_header = null;
                filename = null;
                document_header.Close();
                winword_header.Quit(ref missing, ref missing, ref missing);
                winword_header = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreateDocuments_contents_body()
        {
            try
            {
                string bodyText = "";
                // Create instance for word app 
                Microsoft.Office.Interop.Word.Application winword_body = new Microsoft.Office.Interop.Word.Application();

                // Set animation status for word application
                winword_body.ShowAnimation = false;

                // Set status for word application visible or not 
                winword_body.Visible = false;

                // Create new document 
                Microsoft.Office.Interop.Word.Document document_body = winword_body.Documents.Add();

                // Setup body text
                foreach (string line in keyTerms_contents)
                {
                    bodyText += line;
                }
                // Adding text to document
                document_body.Content.SetRange(0, 0);
                document_body.Content.Text = bodyText + Environment.NewLine;

                // Save document
                //object filename = @"C:\Avanade\" + clientName + "_body.docx";
                var newFilePath = Directory.CreateDirectory(filePathInput + "\\Word\\");
                object filename = newFilePath.FullName + clientName + "_body.docx";
                document_body.SaveAs2(ref filename);
                //document_header = null;
                filename = null;
                document_body.Close();
                winword_body.Quit();
                winword_body = null;

        }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreateDocuments_contents_footer()
        {
            try
            {
                string footerText = "";
                // Create instance for word app 
                Microsoft.Office.Interop.Word.Application winword_footer = new Microsoft.Office.Interop.Word.Application();

                // Set animation status for word application
                winword_footer.ShowAnimation = false;

                // Set status for word application visible or not 
                winword_footer.Visible = false;

                // Create missing variable for missing value 
                object missing = System.Reflection.Missing.Value;

                // Create new document 
                Microsoft.Office.Interop.Word.Document document_footer = winword_footer.Documents.Add();

                // Add footers into the document 
                foreach (Microsoft.Office.Interop.Word.Section wordSection in document_footer.Sections)
                {
                    // Get the footer range and add the footer details 
                    Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                    footerRange.Font.Size = 10;
                    footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    foreach (string line in keyTerms_contents)
                    {
                        footerText += line;
                    }
                    footerRange.Text = keyTerms_properties;
                }

                // Save document
                //object filename = @"C:\Avanade\" + clientName + "_footer.docx";
                var newFilePath = Directory.CreateDirectory(filePathInput + "\\Word\\");
                object filename = newFilePath.FullName + clientName + "_footer.docx";
                document_footer.SaveAs2(ref filename);
                //document_header = null;
                filename = null;
                document_footer.Close();
                winword_footer.Quit(ref missing, ref missing, ref missing);
                winword_footer = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreateDocument_properties_category()
        {
            string bodyText = "";
            // Setup properties text
            foreach (string line in keyTerms_contents)
            {
                bodyText += line;
            }
            try
            {
                // Create instance for word app 
                Microsoft.Office.Interop.Word.Application winword_prop = new Microsoft.Office.Interop.Word.Application();

                // Set animation status for word application
                winword_prop.ShowAnimation = false;

                // Set status for word application visible or not 
                winword_prop.Visible = false;

                // Create missing variable for missing value 
                object missing = System.Reflection.Missing.Value;

                // Create new document 
                Microsoft.Office.Interop.Word.Document doc_prop = winword_prop.Documents.Add();

                doc_prop.BuiltInDocumentProperties["Category"].Value = keyTerms_properties;

                // Save document
                //object filename = @"C:\Avanade\" + clientName + "_propCategory.docx";
                var newFilePath = Directory.CreateDirectory(filePathInput + "\\Word\\");
                object filename = newFilePath.FullName + clientName + "_prop1Category.docx";
                doc_prop.SaveAs2(ref filename);
                filename = null;
                doc_prop.Close();
                winword_prop.Quit();
                winword_prop = null;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreateDocument_properties_keywords()
        {
            string bodyText = "";
            // Setup properties text
            foreach (string line in keyTerms_contents)
            {
                bodyText += line;
            }
            try
            {
                // Create instance for word app 
                Microsoft.Office.Interop.Word.Application winword_prop = new Microsoft.Office.Interop.Word.Application();

                // Set animation status for word application
                winword_prop.ShowAnimation = false;

                // Set status for word application visible or not 
                winword_prop.Visible = false;

                // Create missing variable for missing value 
                object missing = System.Reflection.Missing.Value;

                // Create new document 
                Microsoft.Office.Interop.Word.Document doc_prop = winword_prop.Documents.Add();

                doc_prop.BuiltInDocumentProperties["Keywords"].Value = keyTerms_properties;

                // Save document
                //object filename = @"C:\Avanade\" + clientName + "_propKeywords.docx";
                var newFilePath = Directory.CreateDirectory(filePathInput + "\\Word\\");
                object filename = newFilePath.FullName + clientName + "_prop2Keywords.docx";
                doc_prop.SaveAs2(ref filename);
                filename = null;
                doc_prop.Close();
                winword_prop.Quit();
                winword_prop = null;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CreateExcel_FileNames()
        {
            try
            {
                // Create missing variable for missing value 
                object missing = System.Reflection.Missing.Value;

                // Create instance for word app 
                Microsoft.Office.Interop.Excel.Application winExcel = new Microsoft.Office.Interop.Excel.Application();

                // Set status for word application visible or not 
                winExcel.Visible = false;
                //winExcel.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;

                // SECTION BELOW CREATES 1 DOCUMENT CONTAINING ALL FILE NAMES 
                Microsoft.Office.Interop.Excel.Workbook wb = winExcel.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet ws = wb.Worksheets[1];
                var newFilePath = Directory.CreateDirectory(filePathInput + "\\Excel\\");
                wb.SaveAs(newFilePath.FullName + combinedTitleNames + ".xlsx");
                wb.Close();

                // SECTION BELOW CREATES NEW DOCUMENT FOR EACH FILE NAME 
                //string keyTermStr;
                //foreach (string keyTerm in keyTerms_titles)
                //{
                //    Microsoft.Office.Interop.Excel.Workbook wb = winExcel.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                //    Microsoft.Office.Interop.Excel.Worksheet ws = wb.Worksheets[1];

                //    // Data
                //    keyTermStr = keyTerm.Trim();
                //    var newFilePath = Directory.CreateDirectory(filePathInput + "\\Excel\\");
                //    wb.SaveAs(newFilePath.FullName + keyTermStr + ".xlsx");
                //    keyTermStr = "";

                //    wb.Close();
                //}
                //MessageBox.Show("All Documents should be created successfully!");
                winExcel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreateExcel_contents()
        {
            CreateExcel_contents_body();
            //CreateExcel_contents_header();
            CreateExcel_contents_HeaderFooter();
            //CreateExcel_contents_footer();
            CreateExcel_properties_category();
            CreateExcel_properties_keywords();
        }
        private void CreateExcel_contents_body()
        {
            string lineStr = "";
            try
            {
                // Create instance for word app 
                Microsoft.Office.Interop.Excel.Application winExcel_body = new Microsoft.Office.Interop.Excel.Application();

                // Set status for word application visible or not 
                winExcel_body.Visible = false;
                //winExcel_body.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;

                Microsoft.Office.Interop.Excel.Workbook wb_body = winExcel_body.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet ws_body = wb_body.Worksheets[1];

                // Setup body text
                for (int i=0; i<keyTerms_contents.Count; i++)
                {
                    lineStr = keyTerms_contents[i].Trim();
                    if (!string.IsNullOrEmpty(lineStr))
                    {
                        //ws_body.Range[(i + 1)][1].Value = lineStr;
                        ws_body.Range["B" + (i + 1)].Value = lineStr;
                    }
                    lineStr = "";
                }

                // Save document
                var newFilePath = Directory.CreateDirectory(filePathInput + "\\Excel\\");
                wb_body.SaveAs(newFilePath.FullName + clientName + "_body.xlsx");
                wb_body.Close();
                winExcel_body.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreateExcel_contents_header()
        {
            string lineStr = "";
            int cnt = 0;
            int fileVers = 1;
            try
            {
                // Create instance for word app 
                Microsoft.Office.Interop.Excel.Application winExcel_body = new Microsoft.Office.Interop.Excel.Application();

                // Set status for word application visible or not 
                winExcel_body.Visible = false;
                //winExcel_body.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;
                Microsoft.Office.Interop.Excel.Workbook wb_body = winExcel_body.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet ws_body = wb_body.Worksheets[1];

                // Setup body text
                foreach (string line in keyTerms_contents)
                {
                    if (!string.IsNullOrEmpty(line.Trim()) && (lineStr + line).Length < 250)
                    {
                        lineStr += line;
                    }
                    else if (!string.IsNullOrEmpty(line.Trim()) && (lineStr + line).Length >= 250)
                    {

                    }
                }

                if (fileVers == 1)
                {
                    //ws_body.PageSetup.CenterHeader = lineStr;
                    ws_body.PageSetup.CenterHeader = keyTerms_properties;
                    ws_body.Range["A1"].Value = clientName + "_header";

                    // Save document
                    var newFilePath = Directory.CreateDirectory(filePathInput + "\\Excel\\");
                    wb_body.SaveAs(newFilePath.FullName + clientName + "_header.xlsx");
                    wb_body.Close();
                }
                winExcel_body.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreateExcel_contents_footer()
        {
            string lineStr = "";
            try
            {
                // Create instance for word app 
                Microsoft.Office.Interop.Excel.Application winExcel_body = new Microsoft.Office.Interop.Excel.Application();

                // Set status for word application visible or not 
                winExcel_body.Visible = false;
                //winExcel_body.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;
                Microsoft.Office.Interop.Excel.Workbook wb_body = winExcel_body.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet ws_body = wb_body.Worksheets[1];

                // Setup body text
                foreach (string line in keyTerms_contents)
                {
                    if (!string.IsNullOrEmpty(line.Trim()))
                    {
                        lineStr += line;
                    }
                }
                //ws_body.PageSetup.CenterFooter = lineStr;
                ws_body.PageSetup.CenterFooter = keyTerms_properties;
                ws_body.Range["A1"].Value = clientName + "_footer";

                // Save document
                var newFilePath = Directory.CreateDirectory(filePathInput + "\\Excel\\");
                wb_body.SaveAs(newFilePath.FullName + clientName + "_footer.xlsx");
                wb_body.Close();
                winExcel_body.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreateExcel_properties_category()
        {
            try
            {
                // Create instance for word app 
                Microsoft.Office.Interop.Excel.Application winExcel_prop = new Microsoft.Office.Interop.Excel.Application();

                // Set status for word application visible or not 
                winExcel_prop.Visible = false;

                // Create new document 
                Microsoft.Office.Interop.Excel.Workbook wb_prop = winExcel_prop.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);

                wb_prop.BuiltinDocumentProperties("Category").Value = keyTerms_properties;
                // Save document
                var newFilePath = Directory.CreateDirectory(filePathInput + "\\Excel\\");
                wb_prop.SaveAs(newFilePath.FullName + clientName + "_prop1Category.xlsx");
                wb_prop.Close();
                winExcel_prop.Quit();
                winExcel_prop = null;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreateExcel_properties_keywords()
        {
            try
            {
                // Create instance for word app 
                Microsoft.Office.Interop.Excel.Application winExcel_prop = new Microsoft.Office.Interop.Excel.Application();

                // Set status for word application visible or not 
                winExcel_prop.Visible = false;

                // Create new document 
                Microsoft.Office.Interop.Excel.Workbook wb_prop = winExcel_prop.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);

                wb_prop.BuiltinDocumentProperties("Keywords").Value = keyTerms_properties;

                // Save document
                var newFilePath = Directory.CreateDirectory(filePathInput + "\\Excel\\");
                wb_prop.SaveAs(newFilePath.FullName + clientName + "_prop2Keywords.xlsx");
                wb_prop.Close();
                winExcel_prop.Quit();
                winExcel_prop = null;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CreatePPT_FileNames()
        {
            try
            {
                // Create instance for word app 
                Microsoft.Office.Interop.PowerPoint.Application winPPT = new Microsoft.Office.Interop.PowerPoint.Application();
                // Set status for word application visible or not 
                winPPT.Visible = MsoTriState.msoTrue;

                // SECTION BELOW CREATES 1 DOCUMENT CONTAINING ALL FILE NAMES 
                Microsoft.Office.Interop.PowerPoint.Presentations presSet = winPPT.Presentations;
                Microsoft.Office.Interop.PowerPoint.Presentation pres = presSet.Add();
                Microsoft.Office.Interop.PowerPoint.Slide slide = pres.Slides.AddSlide(1, pres.SlideMaster.CustomLayouts[1]);
                slide.Layout = Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank;

                var newFilePath = Directory.CreateDirectory(filePathInput + "\\PPT\\");
                pres.SaveAs(newFilePath.FullName + combinedTitleNames + ".pptx");
                pres.Close();

                // SECTION BELOW CREATES NEW DOCUMENT FOR EACH FILE NAME 
                //string keyTermStr = "";
                //foreach (string keyTerm in keyTerms_titles)
                //{
                //    Microsoft.Office.Interop.PowerPoint.Presentations presSet = winPPT.Presentations;
                //    Microsoft.Office.Interop.PowerPoint.Presentation pres = presSet.Add();
                //    Microsoft.Office.Interop.PowerPoint.Slide slide = pres.Slides.AddSlide(1, pres.SlideMaster.CustomLayouts[1]);
                //    slide.Layout = Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank;

                //    // Data 
                //    keyTermStr = keyTerm.Trim();
                //    var newFilePath = Directory.CreateDirectory(filePathInput + "\\PPT\\");
                //    pres.SaveAs(newFilePath.FullName + keyTermStr + ".pptx");
                //    pres.Close();
                //}

                winPPT.Quit();


                //Microsoft.Office.Interop.PowerPoint.Slides slides = pres.Slides;

                //// Build Slide #1
                //Microsoft.Office.Interop.PowerPoint.CustomLayout layout = pres.Slides.AddSlide(1, pres.SlideMaster.CustomLayouts[1]);
                ////Microsoft.Office.Interop.PowerPoint.CustomLayout layout = Microsoft.Office.Interop.PowerPoint.CustomLayout;
                //Microsoft.Office.Interop.PowerPoint.Slide slide = slides.AddSlide(1, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank);
                //Microsoft.Office.Interop.PowerPoint.CustomLayouts

                //winExcel.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;

                //string keyTermStr;
                //foreach (string keyTerm in keyTerms_titles)
                //{
                //    Microsoft.Office.Interop.Excel.Workbook wb = winExcel.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                //    Microsoft.Office.Interop.Excel.Worksheet ws = wb.Worksheets[1];

                //    // Data
                //    keyTermStr = keyTerm.Trim();
                //    wb.SaveAs(@"C:\Avanade\" + keyTermStr + ".xlsx");
                //    keyTermStr = "";

                //    wb.Close();
                //}
                //MessageBox.Show("All Documents should be created successfully!");
                //winExcel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreatePPT_contents()
        {
            CreatePPT_contents_body();
            CreatePPT_contents_footer();
            CreatePPT_contents_notes_header();
            CreatePPT_contents_notes_footer();
            CreatePPT_properties_category();
            CreatePPT_properties_keywords();
        }
        private void CreatePPT_contents_body()
        {
            string bodyText = "";
            foreach (string line in keyTerms_contents)
            {
                bodyText += line;
            }
            try
            {
                // Create instance for word app 
                Microsoft.Office.Interop.PowerPoint.Application winPPT = new Microsoft.Office.Interop.PowerPoint.Application();
                // Set status for word application visible or not 
                winPPT.Visible = MsoTriState.msoTrue;

                Microsoft.Office.Interop.PowerPoint.Presentations presSet = winPPT.Presentations;
                Microsoft.Office.Interop.PowerPoint.Presentation pres = presSet.Add();
                Microsoft.Office.Interop.PowerPoint.Slide slide = pres.Slides.AddSlide(1, pres.SlideMaster.CustomLayouts[1]);
                slide.Layout = Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank;
                slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 200).TextFrame.TextRange.Text = bodyText;

                var newFilePath = Directory.CreateDirectory(filePathInput + "\\PPT\\");
                pres.SaveAs(newFilePath.FullName + clientName + "_body.pptx");
                pres.Close();

                winPPT.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreatePPT_contents_footer()
        {
            string bodyText = "";
            foreach (string line in keyTerms_contents)
            {
                if (!string.IsNullOrEmpty(line.Trim()))
                {
                    bodyText += line;
                }
            }
            try
            {
                // Create instance for word app 
                Microsoft.Office.Interop.PowerPoint.Application winPPT = new Microsoft.Office.Interop.PowerPoint.Application();
                // Set status for word application visible or not 
                winPPT.Visible = MsoTriState.msoTrue;

                Microsoft.Office.Interop.PowerPoint.Presentations presSet = winPPT.Presentations;
                Microsoft.Office.Interop.PowerPoint.Presentation pres = presSet.Add();
                Microsoft.Office.Interop.PowerPoint.Slide slide = pres.Slides.AddSlide(1, pres.SlideMaster.CustomLayouts[1]);
                slide.Layout = Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank;
                slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 200).TextFrame.TextRange.Text = clientName + "_footer";

                slide.HeadersFooters.Footer.Visible = MsoTriState.msoTrue;
                slide.HeadersFooters.Footer.Text = keyTerms_properties;

                var newFilePath = Directory.CreateDirectory(filePathInput + "\\PPT\\");
                pres.SaveAs(newFilePath.FullName + clientName + "_footer.pptx");
                pres.Close();

                winPPT.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreatePPT_contents_notes_header()
        {
            string bodyText = "";
            foreach (string line in keyTerms_contents)
            {
                if (!string.IsNullOrEmpty(line.Trim()))
                {
                    bodyText += line;
                }
            }
            try
            {
                // Create instance for word app 
                Microsoft.Office.Interop.PowerPoint.Application winPPT = new Microsoft.Office.Interop.PowerPoint.Application();
                // Set status for word application visible or not 
                winPPT.Visible = MsoTriState.msoTrue;

                Microsoft.Office.Interop.PowerPoint.Presentations presSet = winPPT.Presentations;
                Microsoft.Office.Interop.PowerPoint.Presentation pres = presSet.Add();
                Microsoft.Office.Interop.PowerPoint.Slide slide = pres.Slides.AddSlide(1, pres.SlideMaster.CustomLayouts[1]);
                slide.Layout = Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank;
                slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 200).TextFrame.TextRange.Text = clientName + "_header";

                pres.NotesMaster.HeadersFooters.Header.Text = keyTerms_properties;
                pres.NotesMaster.HeadersFooters.Header.Visible = MsoTriState.msoTrue;

                var newFilePath = Directory.CreateDirectory(filePathInput + "\\PPT\\");
                pres.SaveAs(newFilePath.FullName + clientName + "_notes_header.pptx");
                pres.Close();

                winPPT.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreatePPT_contents_notes_footer()
        {
            string bodyText = "";
            foreach (string line in keyTerms_contents)
            {
                if (!string.IsNullOrEmpty(line.Trim()))
                {
                    bodyText += line;
                }
            }
            try
            {
                // Create instance for word app 
                Microsoft.Office.Interop.PowerPoint.Application winPPT = new Microsoft.Office.Interop.PowerPoint.Application();
                // Set status for word application visible or not 
                winPPT.Visible = MsoTriState.msoTrue;

                Microsoft.Office.Interop.PowerPoint.Presentations presSet = winPPT.Presentations;
                Microsoft.Office.Interop.PowerPoint.Presentation pres = presSet.Add();
                Microsoft.Office.Interop.PowerPoint.Slide slide = pres.Slides.AddSlide(1, pres.SlideMaster.CustomLayouts[1]);
                slide.Layout = Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank;
                slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 200).TextFrame.TextRange.Text = clientName + "_header";

                pres.NotesMaster.HeadersFooters.Footer.Text = keyTerms_properties;
                pres.NotesMaster.HeadersFooters.Footer.Visible = MsoTriState.msoTrue;

                var newFilePath = Directory.CreateDirectory(filePathInput + "\\PPT\\");
                pres.SaveAs(newFilePath.FullName + clientName + "_notes_footer.pptx");
                pres.Close();

                winPPT.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreatePPT_properties_category()
        {
            try
            {
                // Create instance for word app 
                Microsoft.Office.Interop.PowerPoint.Application winPPT = new Microsoft.Office.Interop.PowerPoint.Application();
                // Set status for word application visible or not 
                winPPT.Visible = MsoTriState.msoTrue;

                Microsoft.Office.Interop.PowerPoint.Presentations presSet = winPPT.Presentations;
                Microsoft.Office.Interop.PowerPoint.Presentation pres = presSet.Add();
                Microsoft.Office.Interop.PowerPoint.Slide slide = pres.Slides.AddSlide(1, pres.SlideMaster.CustomLayouts[1]);
                slide.Layout = Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank;
                slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 200).TextFrame.TextRange.Text = clientName + " Properties - Categories";

                pres.BuiltInDocumentProperties["Category"].Value = keyTerms_properties;

                var newFilePath = Directory.CreateDirectory(filePathInput + "\\PPT\\");
                pres.SaveAs(newFilePath.FullName + clientName + "_prop1Category.pptx");
                pres.Close();

                winPPT.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreatePPT_properties_keywords()
        {
            try
            {
                // Create instance for word app 
                Microsoft.Office.Interop.PowerPoint.Application winPPT = new Microsoft.Office.Interop.PowerPoint.Application();
                // Set status for word application visible or not 
                winPPT.Visible = MsoTriState.msoTrue;

                Microsoft.Office.Interop.PowerPoint.Presentations presSet = winPPT.Presentations;
                Microsoft.Office.Interop.PowerPoint.Presentation pres = presSet.Add();
                Microsoft.Office.Interop.PowerPoint.Slide slide = pres.Slides.AddSlide(1, pres.SlideMaster.CustomLayouts[1]);
                slide.Layout = Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank;
                slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 200).TextFrame.TextRange.Text = clientName + " Properties - Keywords";

                pres.BuiltInDocumentProperties["Keywords"].Value = keyTerms_properties;

                var newFilePath = Directory.CreateDirectory(filePathInput + "\\PPT\\");
                pres.SaveAs(newFilePath.FullName + clientName + "_prop2Keywords.pptx");
                pres.Close();

                winPPT.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CreatePDF_all()
        {
            Directory.CreateDirectory(filePathInput + "\\PDF\\");
            // Create a new Microsoft Word application object
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            // C# doesn't have optional arguments so need dummy value 
            object oMissing = System.Reflection.Missing.Value;

            // Get list of Word files in directory
            DirectoryInfo dirInfo = new DirectoryInfo(filePathInput + "\\Word\\");
            FileInfo[] wordFiles = dirInfo.GetFiles("*.docx");

            word.Visible = false;
            word.ScreenUpdating = false;

            foreach (FileInfo wordFile in wordFiles)
            {
                // Cast as Object for word Open method 
                Object filename = (Object)wordFile.FullName;

                // Use the dummy value as a placeholder for optional arguments 
                Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(ref filename, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                doc.Activate();

                //object outputFileName = wordFile.FullName.Replace(".docx", ".pdf");
                //outputFileName = outputFileName.FullName.Replace("Word", "PDF");
                object fileFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;

                string outputFileNameStr = wordFile.FullName.Replace(".docx", ".pdf");
                outputFileNameStr = outputFileNameStr.Replace("\\Word\\", "\\PDF\\");
                object outputFileNameObj = outputFileNameStr;

                // Save document into PDF format 
                doc.SaveAs2(ref outputFileNameObj, ref fileFormat, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                // Close word document, but leave the Word application open 
                // doc has toe be cast to type _Document so that it will find the 
                // correct Close method 
                object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
                ((Microsoft.Office.Interop.Word._Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                doc = null;
            }
            // word has to be cast to type _Application so that it will find 
            // the correct Quit method 
            ((Microsoft.Office.Interop.Word._Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
            word = null;
        }

        private void CreateExcel_contents_HeaderFooter()
        {
            string lineStr = "";
            int cnt = 0;
            int fileVers = 1;

            // Setup body text
            foreach (string line in keyTerms_contents)
            {
                if (!string.IsNullOrEmpty(line.Trim()) && (lineStr + line).Length < 250)
                {
                    //lineStr += DocumentContentsTextBox.GetLineText(line).Trim() + " ";
                    lineStr += line.Trim() + ", ";
                }
                else if (!string.IsNullOrEmpty(line.Trim()) && (lineStr + line).Length >= 250)
                {
                    SaveExcelFile("header", lineStr, fileVers);
                    SaveExcelFile("footer", lineStr, fileVers);
                    lineStr = line.Trim() + ", ";
                    fileVers++;
                }
                cnt++;
            }
            SaveExcelFile("header", lineStr, fileVers);
            SaveExcelFile("footer", lineStr, fileVers);
        }
        private void SaveExcelFile(string fileType, string textContents, int fileVers)
        {
            try
            {
                // Create instance for word app 
                Microsoft.Office.Interop.Excel.Application winExcel_body = new Microsoft.Office.Interop.Excel.Application();

                // Set status for word application visible or not 
                winExcel_body.Visible = false;
                //winExcel_body.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;
                Microsoft.Office.Interop.Excel.Workbook wb_body = winExcel_body.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet ws_body = wb_body.Worksheets[1];

                if (fileType == "header")
                {
                    ws_body.PageSetup.CenterHeader = textContents;
                    ws_body.Range["A1"].Value = clientName + "_header" + fileVers;

                    // Save document
                    var newFilePath = Directory.CreateDirectory(filePathInput + "\\Excel\\");
                    wb_body.SaveAs(newFilePath.FullName + clientName + "_header" + fileVers + ".xlsx");
                }
                else if (fileType == "footer")
                {
                    ws_body.PageSetup.CenterFooter = textContents;
                    ws_body.Range["A1"].Value = clientName + "_footer" + fileVers;

                    // Save document
                    var newFilePath = Directory.CreateDirectory(filePathInput + "\\Excel\\");
                    wb_body.SaveAs(newFilePath.FullName + clientName + "_footer" + fileVers + ".xlsx");
                }

                wb_body.Close();

                winExcel_body.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //private void EmailButton_Click(object sender, EventArgs e)
        //{
        //    Outlook.MailItem mail = Application.CreateItem(
        //        Outlook.OlItemType.olMailItem) as Outlook.MailItem;
        //    mail.Subject = "Test Email";
        //    Outlook.AddressEntry currentUser = Application.Session.CurrentUser.AddressEntry;
        //    if (currentUser.Type == "EX")
        //    {
        //        Outlook.ExchangeUser manager = currentUser.GetExchangeUser().GetExchangeUserManager();
        //        //Add recipient using display name, alias, or smtp address
        //        mail.Recipients.Add(manager.PrimarySmtpAddress);
        //        mail.Recipients.ResolveAll();
        //        mail.Attachments.Add(@"C:\Users\ryan.yoo\Documents\0_DLP\Test\5_Unilever_ZBB\PDF\Unilever_ZBB_body.pdf",
        //            Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
        //        mail.Send();
        //    }

        //}
        private void EmailButton_Click(object sender, EventArgs e)
        {
            // Get List of files to send 
            List<string> listOfEmailFiles = GetFilesToEmail();

            try
            {
                // Create Outlook application
                Outlook.Application oApp = new Outlook.Application();
                // Create new mail item
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                // Set HTMLBody
                // add body of email 
                oMsg.HTMLBody = "Hello, \n\n Please see attached document";
                // add attachment
                String sDisplayName = "MyAttachment";
                int iPosition = (int)oMsg.Body.Length + 1;
                int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                // attach the file 
                Outlook.Attachment oAttach = oMsg.Attachments.Add(@"C:\Users\ryan.yoo\Documents\0_DLP\Test\5_Unilever_ZBB\PDF\Unilever_ZBB_body.pdf",
                    iAttachType, iPosition, sDisplayName);
                // Subject Line
                oMsg.Subject = "Subject Line HERE!!!";
                // Add recipient 
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                // Change recipient in next line if necessary 
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add("ryanyoo1@yahoo.com");
                oRecip.Resolve();
                // Send /
                oMsg.Send();
                // Clean up 
                oRecip = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private List<string> GetFilesToEmail()
        {
            List<string> returnList = new List<string>();

            // Get list of Word files in directory
            DirectoryInfo dirWord = new DirectoryInfo(filePathInput + "\\Word\\");
            FileInfo[] wordFiles = dirWord.GetFiles("*_body.docx");
            returnList.Add(wordFiles[0].FullName);

            // Get list of Excel files in dic
            DirectoryInfo dirExcel= new DirectoryInfo(filePathInput + "\\Excel\\");
            FileInfo[] xlsxFiles = dirExcel.GetFiles("*_body.xlsx");
            returnList.Add(xlsxFiles[0].FullName);

            // Get list of PPT files in dic 
            DirectoryInfo dirPPT = new DirectoryInfo(filePathInput + "\\PPT\\");
            FileInfo[] pptFiles = dirPPT.GetFiles("._body.pptx");
            returnList.Add(pptFiles[0].FullName);

            // Get list of PDF files is dic 
            DirectoryInfo dirPDF = new DirectoryInfo(filePathInput + "\\PDF\\");
            FileInfo[] pdfFiles = dirPDF.GetFiles("._body.pdf");
            returnList.Add(pdfFiles[0].FullName);

            return returnList;
        }
    }
}
