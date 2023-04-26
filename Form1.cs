using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Words;
using static diploma.Form1;
using Aspose.Words.Saving;
using Orientation = Aspose.Words.Orientation;
using Aspose.Words.Tables;

namespace diploma
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        internal static class DBHelper
        {
            public static SqlConnection connection;
            public static DataSet dsLastDataSet;
            public static SqlDataAdapter adapter;
            public static string sLastQuery;
            public static string Connect = @"Data Source = localhost;Initial Catalog = working_hell;Integrated Security = True";
            public static DataTable TryQuery(string query)
            {
                DataTable tmp = new DataTable();
                if (connection == null || adapter == null)
                    TryDeffConnection();
                else
                {
                    try
                    {
                        using (connection = new SqlConnection(Connect))
                        {
                            using (adapter = new SqlDataAdapter(query, connection))
                            {
                                dsLastDataSet = new DataSet();
                                adapter.Fill(dsLastDataSet);
                                sLastQuery = query;
                                Console.WriteLine(sLastQuery);
                            }
                        }
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show("Connection error:\n" + ex.Message);
                    }
                    if (dsLastDataSet.Tables.Count == 0) return tmp; else return dsLastDataSet.Tables[0];
                }
                return null;
            }
            public static void TryDeffConnection()
            {
                try
                {
                    using (connection = new SqlConnection(Connect))
                    {
                        using (adapter = new SqlDataAdapter("select * from Subject;", connection))
                        {
                            dsLastDataSet = new DataSet();
                            adapter.Fill(dsLastDataSet);
                            Console.WriteLine("default connection established");
                        }
                    }
                }
                catch (SqlException ex)
                {
                    MessageBox.Show("Conncetion error:\n" + ex.Message);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //using (WordprocessingDocument document = WordprocessingDocument.Create("D:\\diploma\\diploma.docx", WordprocessingDocumentType.Document))
            //{
            //    // Add a main document part
            //    MainDocumentPart mainPart = document.AddMainDocumentPart();

            //    // Create a new document
            //    Document document12 = new Document();
            //    mainPart.Document = document12;

            //    // Create a new paragraph
            //    Paragraph paragraph = new Paragraph(new Run(new Text("Hello, world!  BEERTODAYDIETOMORROW")));

            //    // Add the paragraph to the document
            //    document12.Body = new Body(paragraph);

            //    // Save the document
            //    mainPart.Document.Save();
            //}
            Document doc = new Document("D:\\diploma\\diploma.docx");
            Section section = doc.Sections[0];
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("\n\n\n");
            section.PageSetup.Orientation = Orientation.Landscape;
            //doc.FirstSection.Body.Remove();
            // Get the first paragraph of the first section
            Paragraph para = doc.FirstSection.Body.FirstParagraph;
            doc.FirstSection.Body.AppendChild(para);
            Run newRun = new Run(doc, "Hello,123 ");
            newRun.Font.Name = "Arial";
            newRun.Font.Size = 12;
            newRun.Font.Bold = true;

            Run refRun = para.Runs[0];
            // Insert text at the beginning of the paragraph
            para.InsertBefore(newRun, refRun);
            Table table = new Table(doc);
            for (int row = 0; row < 6; row++)
            {
                // Add a new row to the table
                Row tableRow = new Row(doc);
                table.Rows.Add(tableRow);

                for (int col = 0; col < 8; col++)
                {
                    // Add a new cell to the row
                    Cell cell = new Cell(doc);
                    tableRow.Cells.Add(cell);

                    // Create a new paragraph and run, and add it to the cell
                    Paragraph cellPara = new Paragraph(doc);
                    //Run cellRun = new Run(doc, "Row " + (row + 1) + ", Col " + (col + 1));
                    //cellPara.AppendChild(cellRun);
                    cell.CellFormat.ClearFormatting();
                    cell.CellFormat.Orientation = TextOrientation.Horizontal;
                    cell.CellFormat.VerticalMerge = CellMerge.None;
                    cell.CellFormat.HorizontalMerge = CellMerge.None;
                    cell.CellFormat.WrapText = true;
                    cell.CellFormat.FitText = true;
                    cell.CellFormat.Borders.LineStyle = LineStyle.Single;
                    cell.CellFormat.Shading.BackgroundPatternColor = Color.White;
                    cell.CellFormat.TopPadding = 5.0;
                    cell.CellFormat.BottomPadding = 5.0;
                    cell.CellFormat.LeftPadding = 5.0;
                    cell.CellFormat.RightPadding = 5.0;
                    cell.AppendChild(cellPara);
                }
            }
            builder.MoveToDocumentEnd();

            // Add the table to the paragraph
            //para.AppendChild(table);
            doc.FirstSection.Body.AppendChild(table);
            // Add the paragraph to the document
            doc.FirstSection.Body.AppendChild(para);
            // Save the modified document
            doc.Save("D:\\diploma\\diploma.docx");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //using (FileStream docStream = new FileStream("D:\\diploma\\diploma.docx", FileMode.Open, FileAccess.Read))
            //{
            //    //Load an existing Word document.
            //    using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Automatic))
            //    {
            //        //Create a new instance of DocIORenderer.
            //        using (DocIORenderer render = new DocIORenderer())
            //        {
            //            //Convert an entire Word document to images.
            //            Stream[] imageStreams = wordDocument.RenderAsImages();
            //            for (int i = 0; i < imageStreams.Length; i++)
            //            {
            //                //Save the stream as file.
            //                using (FileStream fileStreamOutput = File.Create("D:\\diploma\\diploma.jpeg"))
            //                {
            //                    imageStreams[i].CopyTo(fileStreamOutput);
            //                }
            //            }
            //        }
            //    }
            //}
            try
            {
                var doc = new Document("D:\\diploma\\diploma.docx");
                var saveOptions = new Aspose.Words.Saving.ImageSaveOptions(Aspose.Words.SaveFormat.Png);
                saveOptions.PageSet = new PageSet(0);

                // Save the first page of the document as a PNG image
                doc.Save("D:\\diploma\\diploma.png", saveOptions);
            }
            catch
            {
                MessageBox.Show("Закрити ворд");
            }
            

            
        }
    }
}
