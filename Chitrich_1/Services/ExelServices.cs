using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Chitrich_1.Models;
using System.Reflection;
using DocumentFormat.OpenXml.Office2010.ExcelAc;

namespace Chitrich_1.Services
{
    class ExelServices
    {
        public static List<T> Read<T>(string filePath) where T : BaseClass, new()
        {
            List<People> peoples = new List<People>();
            if (SpreadsheetDocument.Open(filePath, false) == null)
            {
                throw new Exception("spreadsheetDocument is null");
            }
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                if (spreadsheetDocument.WorkbookPart == null)
                {
                    throw new Exception();
                }
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault()!;
                if (sheet.Id == null)
                {
                    throw new Exception("Sheet not found");
                }
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                int rowNum = 0;
                List<T> returnedList = new List<T>();
                foreach (Row row in sheetData.Elements<Row>())
                {
                    if (rowNum > 0)
                    {
                        var returnedObj = new T();
                        returnedObj.OdjFromRow(row, spreadsheetDocument);

                        returnedList.Add(returnedObj);
                    }
                    rowNum++;
                }
                return returnedList;
            }
        }

        public static void Save(List<People> peoples)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(@"C:\Users\taras\source\repos\Chitrich_1\Chitrich_1\Files\output.xlsx", SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                Sheets sheets = spreadsheetDocument.WorkbookPart!.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };
                sheets.Append(sheet);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                if (sheetData == null)
                {
                    throw new Exception();
                }

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
                    choise = int.Parse(Console.ReadLine() ?? "0");
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
                            headerRow.Append(CreateTextCell("Id"));
                            break;
                        case 2:
                            headerRow.Append(CreateTextCell("Name"));
                            break;
                        case 3:
                            headerRow.Append(CreateTextCell("Age"));
                            break;
                        case 4:
                            headerRow.Append(CreateTextCell("Salary"));
                            break;
                        case 5:
                            headerRow.Append(CreateTextCell("Department"));
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
                                dataRow.Append(CreateIntCell(people.Id));
                                break;
                            case 2:
                                if (people.Name != null)
                                {
                                    dataRow.Append(CreateTextCell(people.Name));
                                    break;
                                }
                                else
                                {
                                    throw new Exception();
                                }
                            case 3:
                                dataRow.Append(CreateIntCell(people.Age));
                                break;
                            case 4:
                                dataRow.Append(CreateIntCell(people.Salary));
                                break;
                            case 5:
                                if (people.Department != null)
                                {
                                    dataRow.Append(CreateTextCell(people.Department));
                                    break;
                                }else
                                {
                                    throw new Exception();
                                }
                                
                        }
                    }
                    sheetData.AppendChild(dataRow);
                }
                #endregion

                workbookPart.Workbook.Save();
            }
        }

        private static Cell CreateTextCell(string cellValue)
        {
            Cell cell = new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue(cellValue)
            };
            return cell;
        }

        private static Cell CreateIntCell( int cellValue)
        {
            Cell cell = new Cell()
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(cellValue)
            };
            return cell;
        }

        
    }
}