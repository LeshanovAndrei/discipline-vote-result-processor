using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WindowsFormsApp4
{
    class FileWriter
    {
        private WorkbookPart workbookpart;
        private SpreadsheetDocument spreadsheetDocument;
        private string filepath;

        public FileWriter(string filepath)
        {
            this.filepath = filepath;
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            spreadsheetDocument = SpreadsheetDocument.
                 Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            //Sheet sheet = new Sheet()
            //{
            //    Id = spreadsheetDocument.WorkbookPart.
            //    GetIdOfPart(worksheetPart),
            //    SheetId = 1,
            //    Name = "mySheet"
            //};
            //sheets.Append(sheet);

            workbookpart.Workbook.Save();

            //// Close the document.
            //spreadsheetDocument.Close();
        }
        //Добавление листа
        public void InsertWorksheet(string name)
        {

            //using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(docName, true))
            //{

            WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart);

            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string sheetName = name;


            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookpart.Workbook.Save();



            //}
        }

        private WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
        {
            string relId = workbookPart.Workbook.Descendants<Sheet>().First(s => sheetName.Equals(s.Name)).Id;
            return (WorksheetPart)workbookPart.GetPartById(relId);
        }

        public void SetCellValue(string columnName, uint rowIndex, string val, string sheetName)
        {
            WorksheetPart worksheetPart = GetWorksheetPart(workbookpart, sheetName);

            // Insert cell A1 into the new worksheet.
            Cell cell = InsertCellInWorksheet(columnName, rowIndex, worksheetPart);

            // Set the value of cell A1.
            cell.CellValue = new CellValue(val);
            cell.DataType = new EnumValue<CellValues>(DataTypeDef(val));

            // Save the new worksheet.
            worksheetPart.Worksheet.Save();
        }

        private CellValues DataTypeDef(string val)
        {
            for (int i = 0; i < val.Length; i++)
            {
                if (val[i] < 48 || val[i] > 57)
                {
                    return CellValues.String;
                }
            }
            return CellValues.Number;
        }

        private Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        public void Close()
        {
            spreadsheetDocument.Close();
        }

        ~FileWriter()
        {
            
        }
    }
}
