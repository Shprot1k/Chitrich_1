using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Chitrich_1.Models;

namespace Chitrich_1.Services
{
    class ExelServices
    {
        public static List<People> Read(string fileName)
        {
            List<People> peoples = new List<People>();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                if (sheet == null)
                {
                    throw new Exception("Sheet not found");
                }
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                int rowNum = 0;
                foreach (Row row in sheetData.Elements<Row>())
                {
                    People p = new People();
                    int cellNum = 0;
                    if (rowNum > 0)
                    {
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            string cellValue = GetCellValue(spreadsheetDocument, cell);
                            switch (cellNum)
                            {
                                case 0:
                                    p.Id = int.Parse(cellValue);
                                    cellNum++;
                                    break;
                                case 1:
                                    p.Name = cellValue;
                                    cellNum++;
                                    break;
                                case 2:
                                    p.Age = int.Parse(cellValue);
                                    cellNum++;
                                    break;
                                case 3:
                                    p.Salary = int.Parse(cellValue);
                                    cellNum++;
                                    break;
                                case 4:
                                    p.Department = cellValue;
                                    cellNum++;
                                    break;
                                default:
                                    cellNum++; break;
                            }
                        }
                        peoples.Add(p);
                    }
                    rowNum++;
                }
            }
            return peoples;
        }
        
        public static void Save(List<People> peoples)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(@"C:\Users\taras\source\repos\Chitrich_1\Chitrich_1\Files\output.xlsx", SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };
                sheets.Append(sheet);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                int choise = 1;
                List<int> choises = new List<int>();

                while (choise != 0)
                {
                    Console.Write("Select the columns to save\n"
                        + "1 - Id\n"
                        + "2 - Name\n"
                        + "3 - Age\n"
                        + "4 - Salary\n"
                        + "5 - Department\n"
                        + "Else - exit\n"
                        + "Your choise: ");
                    choise = int.Parse(Console.ReadLine());
                    Console.WriteLine();
                    if (choise >= 1 && choise <= 5)
                    {
                        choises.Add(choise);
                    }
                    else
                    {
                        choise = 0;
                    }
                }
                #region header

                Row headerRow = new Row();
                int i = 1;
                foreach (int chois in choises)
                {
                    switch (chois)
                    {
                        case 1:
                            headerRow.Append(CreateTextCell(CellNum(i), "Id"));
                            break;
                        case 2:
                            headerRow.Append(CreateTextCell(CellNum(i), "Name"));
                            break;
                        case 3:
                            headerRow.Append(CreateTextCell(CellNum(i), "Age"));
                            break;
                        case 4:
                            headerRow.Append(CreateTextCell(CellNum(i), "Salary"));
                            break;
                        case 5:
                            headerRow.Append(CreateTextCell(CellNum(i), "Department"));
                            break;
                    }
                    i++;
                }
                #endregion
                #region calls
                sheetData.AppendChild(headerRow);
                foreach (People people in peoples)
                {
                    Row dataRow = new Row();
                    foreach (int chois in choises)
                    {
                        switch (chois)
                        {
                            case 1:
                                dataRow.Append(CreateIntCell(null, people.Id));
                                break;
                            case 2:
                                dataRow.Append(CreateTextCell(null, people.Name));
                                break;
                            case 3:
                                dataRow.Append(CreateIntCell(null, people.Age));
                                break;
                            case 4:
                                dataRow.Append(CreateIntCell(null, people.Salary));
                                break;
                            case 5:
                                dataRow.Append(CreateTextCell(null, people.Department));
                                break;
                        }
                    }
                    sheetData.AppendChild(dataRow);
                }
                #endregion

                workbookPart.Workbook.Save();
            }
        }

        private static Cell CreateTextCell(string cellReference, string cellValue)
        {
            Cell cell = new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue(cellValue)
            };

            if (!string.IsNullOrEmpty(cellReference))
            {
                cell.CellReference = cellReference;
            }

            return cell;
        }

        private static Cell CreateIntCell(string cellReference, int cellValue)
        {
            Cell cell = new Cell()
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(cellValue)
            };

            if (!string.IsNullOrEmpty(cellReference))
            {
                cell.CellReference = cellReference;
            }

            return cell;
        }

        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            string value = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                SharedStringTablePart stringTable = document.WorkbookPart.SharedStringTablePart;
                return stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
            }
            else if (cell.DataType != null && cell.DataType.Value == CellValues.Boolean)
            {
                return value == "0" ? "FALSE" : "TRUE";
            }
            else
            {
                return value;
            }
        }

        private static string CellNum(int i)
        {
            switch (i)
            {
                case 1:
                    return "A1";
                case 2:
                    return "B1";
                case 3:
                    return "C1";
                case 4:
                    return "D1";
                case 5:
                    return "E1";
                default:
                    return "Eror";
            }
        }
    }
}
