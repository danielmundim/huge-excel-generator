using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace HugeExcelGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            string outputPath = "path_to_file_here";
            int iterations = 100000;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create(outputPath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                //create workbook part
                WorkbookPart wbp = spreadsheet.AddWorkbookPart();
                wbp.Workbook = new Workbook();
                Sheets sheets = wbp.Workbook.AppendChild<Sheets>(new Sheets());

                //create worksheet part, and add it to the sheets collection in workbook
                WorksheetPart wsp = wbp.AddNewPart<WorksheetPart>();
                Sheet sheet = new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(wsp), SheetId = 1, Name = "Title" };
                sheets.Append(sheet);

                OpenXmlWriter writer = OpenXmlWriter.Create(wsp);
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());
                //
                writer.WriteStartElement(new Row());
                writer.WriteElement(new Cell { CellValue = new CellValue("Test Column 1"), DataType = CellValues.String });
                writer.WriteElement(new Cell { CellValue = new CellValue("Test Column 2"), DataType = CellValues.String });
                writer.WriteElement(new Cell { CellValue = new CellValue("Test Column 3"), DataType = CellValues.String });

                writer.WriteEndElement(); //end of Row

                for (int i = 1; i <= iterations; i++)
                {
                    writer.WriteStartElement(new Row());
                    writer.WriteElement(new Cell { CellValue = new CellValue(i.ToString()), DataType = CellValues.String });
                    writer.WriteElement(new Cell { CellValue = new CellValue("This is a test"), DataType = CellValues.String });
                    writer.WriteElement(new Cell { CellValue = new CellValue(DateTime.Now.ToShortDateString()), DataType = CellValues.String });
                    writer.WriteEndElement(); //end of Row
                    //
                    Console.Clear();
                    Console.WriteLine(string.Format("{0} of {1} processed rows in {2}", i.ToString(), iterations.ToString(), stopwatch.Elapsed.ToString()));
                }

                writer.WriteEndElement(); //end of SheetData
                writer.WriteEndElement(); //end of worksheet
                writer.Close();
            }
        }
    }
}
