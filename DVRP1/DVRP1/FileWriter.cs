using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FileProcessor
{
    class FileWriter
    {
        private WorkbookPart workbookpart;
        private SpreadsheetDocument spreadsheetDocument;
        private string filepath;
        List<SheetData> OutSheets;
        List<string> OutSheetsNames;

        //public
        public FileWriter(string filepath)
        {
            OutSheets = new List<SheetData>();
            OutSheetsNames = new List<string>();
            this.filepath = filepath;
            spreadsheetDocument = SpreadsheetDocument.
                 Create(filepath, SpreadsheetDocumentType.Workbook);

            workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            workbookpart.Workbook.Save();
        }

        //add sheet to doc
        public void InsertWorksheet(string name)
        {
            OutSheets.Add(new SheetData());
            OutSheetsNames.Add(name);
        }

        //write string into cell
        public void SetCellValue(string columnName, uint rowIndex, string val, string sheetName)
        {
            // Insert cell A1 into the new worksheet.
            Cell cell = InsertCellInWorksheet(columnName, rowIndex, NumOfSheet(sheetName));
            // Set the value of cell A1.
            cell.CellValue = new CellValue(val);
            cell.DataType = new EnumValue<CellValues>(DataTypeDef(val));
        }

        //close and save file after writing
        public void CloseAndExport()
        {
            for (int i = 0; i < OutSheetsNames.Count; i++)
            {
                WorksheetPart newWorksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(OutSheets[i]);

                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                string relationshipId = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart);

                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                string sheetName = OutSheetsNames[i];

                Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                sheets.Append(sheet);
            }
            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();
        }

        public string NameOfSheet(int num)
        {
            return OutSheetsNames[num];
        }

        public int NumOfSheet(string name)
        {
            for (int i = 0; i < OutSheetsNames.Count; i++)
            {
                if (OutSheetsNames[i] == name)
                    return i;
            }
            return -1;
        }

        public int LetterToInt(string letter)
        {
            byte[] asciiBytes = Encoding.ASCII.GetBytes(letter);
            return asciiBytes[0] - 64;
        }
        public string IntToLetter(int intValue)
        {
            return ((char)(intValue + 64)).ToString();
        }

        //private
        private WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
        {
            string relId = workbookPart.Workbook.Descendants<Sheet>().First(s => sheetName.Equals(s.Name)).Id;
            return (WorksheetPart)workbookPart.GetPartById(relId);
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

        private Cell InsertCellInWorksheet(string columnName, uint rowIndex, int sheetNum)
        {
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (OutSheets[sheetNum].Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = OutSheets[sheetNum].Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                OutSheets[sheetNum].Append(row);
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
                return newCell;
            }
        }
        ~FileWriter()
        {

        }

    }
}
