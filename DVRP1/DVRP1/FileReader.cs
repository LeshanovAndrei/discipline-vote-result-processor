using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FileProcessor
{
    class FileReader
    {
        private string filepath;
        WorkbookPart workbookPart;
        SpreadsheetDocument spreadsheetDocument;
        List<SheetData> sheets;

        //Public
        //open file for reading by name
        public FileReader(string filepath)
        {
            sheets = new List<SheetData>();
            this.filepath = filepath;
            spreadsheetDocument = SpreadsheetDocument.Open(filepath, false);
            workbookPart = spreadsheetDocument.WorkbookPart;
            int numOfSheets = SheetNumber();
            for (int i = 0; i < numOfSheets; i++)
            {
                Sheet mysheet = (Sheet)workbookPart.Workbook.Sheets.ChildElements.GetItem(i);
                Worksheet worksheet = ((WorksheetPart)workbookPart.GetPartById(mysheet.Id)).Worksheet;
                SheetData currentSheet = (SheetData)worksheet.GetFirstChild<SheetData>();
                sheets.Add(currentSheet);
            }
        }

        public string GetCellValue(string columnName, uint rowIndex, int sheetName)
        {
            string value = "";
            if (rowIndex - 1 >= sheets[sheetName].ChildElements.Count)
            {
                value = GetCellValueOldVers(columnName, rowIndex, NameOfSheet(sheetName));
                if (value == null)
                {
                    value = "";
                }
                return value;
            }
            Row row = (Row)sheets[sheetName].ChildElements.GetItem((int)rowIndex - 1);
            if (LetterToInt(columnName) - 1 >= row.ChildElements.Count)
            {
                value = GetCellValueOldVers(columnName, rowIndex, NameOfSheet(sheetName));
                if (value == null)
                {
                    value = "";
                }
                return value;
            }
            Cell cell = (Cell)row.ChildElements.GetItem(LetterToInt(columnName) - 1);
            if (cell.CellReference != columnName + rowIndex.ToString())
            {
                value = GetCellValueOldVers(columnName, rowIndex, NameOfSheet(sheetName));
                if (value == null)
                {
                    value = "";
                }
                return value;
            }
            else
                value = cell.InnerText;

            if (cell.DataType != null)
            {
                switch (cell.DataType.Value)
                {
                    case CellValues.SharedString:
                        var stringTable =
                            workbookPart.GetPartsOfType<SharedStringTablePart>()
                            .FirstOrDefault();
                        if (stringTable != null)
                        {
                            value =
                                stringTable.SharedStringTable
                                .ElementAt(int.Parse(value)).InnerText;
                        }
                        break;
                    case CellValues.Boolean:
                        switch (value)
                        {
                            case "0":
                                value = "FALSE";
                                break;
                            default:
                                value = "TRUE";
                                break;
                        }
                        break;
                }
            }
            if (value == null)
            {
                value = "";
            }
            return value;
        }

        public string NameOfSheet(int num)
        {
            return workbookPart.Workbook.Descendants<Sheet>().ToList()[num].Name;
        }

        //num of sheets
        public int SheetNumber()
        {
            return workbookPart.Workbook.Descendants<Sheet>().ToList().Count;
        }

        //close file after reading
        public void Close()
        {
            sheets.Clear();
            spreadsheetDocument.Close();
        }

        //shouldn't be here don't pay attention
        public string FacultyFromFileName()
        {
            return filepath;
        }
        //to search by columns
        public int LetterToInt(string letter)
        {
            byte[] asciiBytes = Encoding.ASCII.GetBytes(letter);
            return asciiBytes[0] - 64;
        }
        //to search by columns
        public string IntToLetter(int intValue)
        {
            return ((char)(intValue + 64)).ToString();
        }

        //private
        private string GetCellValueOldVers(string columnName, uint rowIndex, string sheetName)
        {
            string value = null;
            Sheet theSheet = workbookPart.Workbook.Descendants<Sheet>().
              Where(s => s.Name == sheetName).FirstOrDefault();
            if (theSheet == null)
            {
                throw new ArgumentException(sheetName + " not found");
            }
            WorksheetPart wsPart =
                (WorksheetPart)(workbookPart.GetPartById(theSheet.Id));
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().
             Where(c => c.CellReference == (columnName + rowIndex.ToString())).FirstOrDefault();
            if (theCell != null)
            {
                value = theCell.InnerText;
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            var stringTable =
                                workbookPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();
                            if (stringTable != null)
                            {
                                value =
                                    stringTable.SharedStringTable
                                    .ElementAt(int.Parse(value)).InnerText;
                            }
                            break;
                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }
            return value;
        }

    }

}
