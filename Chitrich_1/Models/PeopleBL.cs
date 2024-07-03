using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Chitrich_1.Models
{
    internal class PeopleBL
    {
        public static List<People> Read(string fileName)
        {
            List < People > peoples = new List < People >();
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

        public static void Print(List<People> peoples)
        {
            Console.WriteLine();
            foreach (People people in peoples)
            {
                Console.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}", people.Id, people.Name, people.Age, people.Salary, people.Department);
            }
        }

        public static List<People> Sort(List<People> peoples)
        {
            Console.Write("\nSelect a sorting method:\n"
                + "1 - By Id\n"
                + "2 - By Name (dont work)\n"
                + "3 - By Age\n"
                + "4- By Salare\n"
                + "5 - By Department (dont work)\n"
                + "0 - Don't sort\n"
                + "Your choice: ");
            int sortOption = int.Parse(Console.ReadLine());
            switch (sortOption)
            {
                case 0: // Don't sort
                    return peoples;

                case 1: // Id
                    for (int i = 0; i < peoples.Count; i++)
                    {
                        for (int j = 0; j < peoples.Count; j++)
                        {
                            if (peoples[i].Id < peoples[j].Id)
                            {
                                var temp = peoples[i];
                                peoples[i] = peoples[j];
                                peoples[j] = temp;
                            }
                        }
                    }
                    return peoples;

                case 2: // By Name (dont work)
                    return peoples;

                case 3: // By Age
                    for (int i = 0; i < peoples.Count; i++)
                    {
                        for (int j = 0; j < peoples.Count; j++)
                        {
                            if (peoples[i].Age < peoples[j].Age)
                            {
                                var temp = peoples[i];
                                peoples[i] = peoples[j];
                                peoples[j] = temp;
                            }
                        }
                    }
                    return peoples;

                case 4: // By Salare
                    for (int i = 0; i < peoples.Count; i++)
                    {
                        for (int j = 0; j < peoples.Count; j++)
                        {
                            if (peoples[i].Salary < peoples[j].Salary)
                            {
                                var temp = peoples[i];
                                peoples[i] = peoples[j];
                                peoples[j] = temp;
                            }
                        }
                    }
                    return peoples;

                case 5: //By Department (dont work)
                    return peoples;
                default: 
                    return peoples;


            }
        }

        public static void Save(List<People> peoples)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create("output.xlsx", SpreadsheetDocumentType.Workbook))
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

                workbookPart.Workbook.Save();
            }
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