using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

        public FileReader(string filepath)
        {
            this.filepath = filepath;
            try
            {
                spreadsheetDocument = SpreadsheetDocument.Open(filepath, false);
            }
            catch (Exception)
            {
                MessageBox.Show("File is occupied by another process! Close it!", "ERROR");


            }
            finally
            {
                spreadsheetDocument = SpreadsheetDocument.Open(filepath, false);
                workbookPart = spreadsheetDocument.WorkbookPart;
            }



        }

        public string GetCellValue(string columnName, uint rowIndex, string sheetName)
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

                // If the cell represents an integer number, you are done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and 
                // Booleans individually. For shared strings, the code 
                // looks up the corresponding value in the shared string 
                // table. For Booleans, the code converts the value into 
                // the words TRUE or FALSE.
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:

                            // For shared strings, look up the value in the
                            // shared strings table.
                            var stringTable =
                                workbookPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();

                            // If the shared string table is missing, something 
                            // is wrong. Return the index that is in
                            // the cell. Otherwise, look up the correct text in 
                            // the table.
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

        public string NameOfSheet(int num)
        {
            return workbookPart.Workbook.Descendants<Sheet>().ToList()[num].Name;
        }

        public void Close()
        {
            spreadsheetDocument.Close();
        }
        public string FacultyFromFileName()
        {
            return filepath;
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
    }
}
