/*
 * MIT License
 * 
 * Copyright(c) 2020 ShmoopySoft (Pty) Ltd
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and 
 * associated documentation files (the "Software"), to deal in the Software without restriction, including 
 * without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
 * copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the 
 * following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all copies or substantial 
 * portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT 
 * LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO 
 * EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER 
 * IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE 
 * USE OR OTHER DEALINGS IN THE SOFTWARE.
*/

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ShmoopySoft_Generate_Open_XML_Example
{
    /// <summary>
    /// The Program class's responsibility is to provide an entry point for the application.
    /// </summary>
    class Program
    {
        /// <summary>
        /// C# applications have an entry point called the Main Method. 
        /// It is the first method that gets invoked when an application starts.
        /// </summary>
        /// <param name="args">Command line arguments as string type parameters</param>
        static void Main(string[] args)
        {
            //// String variables to store Report properties
            string myReportName = "My Report";
            string myReportBlurb = "This is an example to show how easy it is to create a Word document using C# and OpenXML.";
            string reportWordFilename = @"C:\Temp\My Report.docx";
            string reportExcelFilename = @"C:\Temp\My Report.xlsx";

            //// Create a sample Report DataTable to export
            DataTable reportDataTable = new DataTable();
            reportDataTable.TableName = "MyReport";

            //// Insert 3 columns
            reportDataTable.Columns.Add("Name", typeof(string));
            reportDataTable.Columns.Add("Number", typeof(decimal));
            reportDataTable.Columns.Add("Date", typeof(DateTime));

            //// Insert 10 Rows
            for (int i = 1; i < 11; i++)
            {
                DataRow newRow = reportDataTable.NewRow();
                newRow["Name"] = "Sample Data " + i.ToString();
                newRow["Number"] = i * 100;
                newRow["Date"] = DateTime.Now.AddMonths(-(i));
                reportDataTable.Rows.Add(newRow);
            }

            try
            {
                //// Generate the Word Document
                ExportReportToWord(reportWordFilename, myReportName, myReportBlurb, reportDataTable);

                //// Generate the Excel Spreadsheet
                ExportDataTableToExcel(reportDataTable, reportExcelFilename);

                // Display a confirmation.
                Console.WriteLine("The reports were successfully generated :-)");
                Console.Write(Environment.NewLine);
                Console.WriteLine("Word Report filename: " + reportWordFilename);
                Console.WriteLine("Excel Report filename: " + reportExcelFilename);
            }
            catch (Exception ex)
            {
                // Display an error.
                Console.WriteLine("Failed to generate the reports :-(");
                Console.Write(Environment.NewLine);
                Console.WriteLine(ex.ToString());
            }

            Console.Write(Environment.NewLine);
            Console.WriteLine("Press any key to end...");
            Console.ReadKey(true);
        }

        /// <summary>
        /// Convert a DataTable to a word processing Table to insert into a document.
        /// </summary>
        /// <param name="dataTable"></param>
        /// <returns>The Table</returns>
        public static Table ConvertDataTableToWordTable(DataTable dataTable)
        {
            //// Create a new table
            Table table = new Table();

            //// Create the table properties
            TableProperties tblProperties = new TableProperties();

            TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
            TableStyle tableStyle = new TableStyle() { Val = "TableGrid" };
            tblProperties.Append(tableStyle, tableWidth);

            //// Create Table Borders
            TableBorders tblBorders = new TableBorders();

            TopBorder topBorder = new TopBorder();
            topBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
            topBorder.Color = "000000";
            tblBorders.AppendChild(topBorder);

            BottomBorder bottomBorder = new BottomBorder();
            bottomBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
            bottomBorder.Color = "000000";
            tblBorders.AppendChild(bottomBorder);

            RightBorder rightBorder = new RightBorder();
            rightBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
            rightBorder.Color = "000000";
            tblBorders.AppendChild(rightBorder);

            LeftBorder leftBorder = new LeftBorder();
            leftBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
            leftBorder.Color = "000000";
            tblBorders.AppendChild(leftBorder);

            InsideHorizontalBorder insideHBorder = new InsideHorizontalBorder();
            insideHBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
            insideHBorder.Color = "000000";
            tblBorders.AppendChild(insideHBorder);

            InsideVerticalBorder insideVBorder = new InsideVerticalBorder();
            insideVBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
            insideVBorder.Color = "000000";
            tblBorders.AppendChild(insideVBorder);

            //// Add the table borders to the properties
            tblProperties.AppendChild(tblBorders);

            //// Add the table properties to the table
            table.AppendChild(tblProperties);

            //// Get the number of columns and rows in the DataTable
            int iColCount = dataTable.Columns.Count;
            int iRowCount = dataTable.Rows.Count;

            //// If there are rows....
            if (iRowCount > 0)
            {
                //// Create a new Table row
                TableRow trColumns = new TableRow();

                for (int i = 0; i < iColCount; i++)
                {
                    //// Add a cell for each column
                    TableCell tc = new TableCell(new Paragraph(new Run(new Text(dataTable.Columns[i].ToString()))));

                    //// Add the cells to the row
                    trColumns.Append(tc);
                }

                //// Add the row to the table
                table.AppendChild(trColumns);

                foreach (DataRow dr in dataTable.Rows)
                {
                    //// Create a new Table row
                    TableRow trDataRow = new TableRow();

                    for (int i = 0; i < iColCount; i++)
                    {
                        if (!Convert.IsDBNull(dr[i]))
                        {
                            //// Add a cell for each column in the row
                            TableCell tc = new TableCell(new Paragraph(new Run(new Text(dr[i].ToString()))));

                            //// Add the cells to the row
                            trDataRow.Append(tc);
                        }
                        else
                        {
                            //// Add an empty cell for each column in the row
                            TableCell tc = new TableCell(new Paragraph(new Run(new Text(""))));

                            //// Add the cells to the row
                            trDataRow.Append(tc);
                        }
                    }

                    //// Add the rows to the table
                    table.AppendChild(trDataRow);
                }
            }

            //// Return the Table
            return table;
        }

        /// <summary>
        /// Creates a word processing Paragraph to insert into a document.
        /// </summary>
        /// <param name="paragraphText">The paragraph text</param>
        /// <param name="textSize">The paragraph font size</param>
        /// <returns>The Paragraph</returns>
        public static Paragraph CreateOpenXmlParagraph(string paragraphText, int textSize)
        {
            //// Create new RunProperties
            RunProperties runProperties = new RunProperties();

            //// Create new FontSize to set the font size
            FontSize fontSize = new FontSize() { Val = textSize.ToString() };

            //// Append the font size to the RunProperties
            runProperties.Append(fontSize);

            //// Create a new Paragraph
            Paragraph paragraph = new Paragraph();

            //// Create a new Run
            Run run = new Run();

            //// Create new Text
            Text text = new Text(paragraphText);

            //// Append the run properties to the new run
            run.Append(runProperties);

            //// Append the text to the new run
            run.Append(text);

            //// Append the new run to the paragraph
            paragraph.Append(run);

            //// Return the paragraph
            return paragraph;
        }

        /// <summary>
        /// Creates a word processing Paragraph formatted as a heading to insert into a document. 
        /// The text is automatically set to bold.
        /// </summary>
        /// <param name="headingText">The heading text</param>
        /// <param name="textSize">The heading font size</param>
        /// <returns>The heading formatted as a Paragraph</returns>
        public static Paragraph CreateOpenXmlHeading(string headingText, int textSize)
        {
            //// Create new RunProperties
            RunProperties runProperties = new RunProperties();

            //// Create new FontSize to set the font size
            FontSize fontSize = new FontSize() { Val = textSize.ToString() };

            //// Create new Bold to set the font as bold
            Bold bold = new Bold();

            //// Append the Bold to the RunProperties
            runProperties.Append(bold);

            //// Append the FontSize to the RunProperties
            runProperties.Append(fontSize);

            //// Create a new Paragraph
            Paragraph paragraph = new Paragraph();

            //// Create a new Run
            Run run = new Run();

            //// Create new Text
            Text text = new Text(headingText);

            //// Append the run properties to the new run
            run.Append(runProperties);

            //// Append the text to the new run
            run.Append(text);

            //// Append the new run to the paragraph
            paragraph.Append(run);

            //// Return the paragraph
            return paragraph;
        }

        /// <summary>
        /// Exports an example report to Word (docx) format.
        /// </summary>
        /// <param name="reportPath">The full path and filename of the docx file to create</param>
        /// <param name="reportName">The name (heading) of the report</param>
        /// <param name="reportBlurb">A blurb (paragraph) of the report</param>
        /// <param name="reportData">A DataTable to convert to a table and insert into the document</param>
        public static void ExportReportToWord(string reportPath, string reportName, string reportBlurb, DataTable reportData)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(reportPath, WordprocessingDocumentType.Document))
            {
                //// Defines the MainDocumentPart            
                MainDocumentPart mainDocumentPart = doc.AddMainDocumentPart();
                mainDocumentPart.Document = new Document();

                //// Add a body to the document main part
                Body body = mainDocumentPart.Document.AppendChild(new Body());

                //// Add a heading to the body
                body.Append(CreateOpenXmlHeading(reportName, 36));

                //// Add a paragraph to the body
                body.Append(CreateOpenXmlParagraph(reportBlurb, 24));

                if (reportData.Rows.Count > 0)
                {
                    //// Add a table to the body
                    body.AppendChild(ConvertDataTableToWordTable(reportData));

                    //// Add a paragraph to the body
                    body.Append(CreateOpenXmlParagraph(" ", 14));
                }
                else
                {
                    //// Add a paragraph to the body
                    body.Append(CreateOpenXmlParagraph("There is no report data to display.", 20));
                }

                //// Save the document
                mainDocumentPart.Document.Save();
            }
        }

        /// <summary>
        /// Exports a DataTable to an Excel (xlsx) spreadsheet format.
        /// </summary>
        /// <param name="dataTable">The DataTable to convert</param>
        /// <param name="excelPath">The full path and filename of the Excel spreadsheet to create</param>
        public static void ExportDataTableToExcel(DataTable dataTable, string excelPath)
        {
            using (var workbook = SpreadsheetDocument.Create(excelPath, SpreadsheetDocumentType.Workbook))
            {
                //// Add a new WorkbookPart
                var workbookPart = workbook.AddWorkbookPart();

                //// Add a new Workbook
                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook
                {
                    Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets()
                };

                var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();

                sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = 
                    workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();

                string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                uint sheetId = 1;

                if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                {
                    sheetId =
                        sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Max(s => s.SheetId.Value) + 1;
                }

                DocumentFormat.OpenXml.Spreadsheet.Sheet sheet =
                    new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = dataTable.TableName };

                sheets.Append(sheet);

                DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                foreach (DataColumn column in dataTable.Columns)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
                    {
                        DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                        CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName)
                    };
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (DataRow dsrow in dataTable.Rows)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
                        {
                            DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[i].ToString())
                        };

                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }
            }
        }
    }
}
