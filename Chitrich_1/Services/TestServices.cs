/*
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

namespace Chitrich_1.Services
{
    internal class TestServices
    {
        public static List<T> Read<T>(string fileName) where T : class
        {
            if (SpreadsheetDocument.Open(fileName, false) != null)
            {
                List<People> peoples = new List<People>();
                List<T> ret = new List<T>();
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
                {
                    if (spreadsheetDocument.WorkbookPart != null)
                    {
                        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                        if (workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault() != null)
                        {
                            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                            if (sheet == null)
                            {
                                
                            }
                            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                            if (typeof(T).GetConstructor(new[] { typeof(List<string>) }) != null)
                            {
                                ConstructorInfo constructor = typeof(T).GetConstructor(new[] { typeof(List<string>) });

                                int rowNum = 0;
                                foreach (Row row in sheetData.Elements<Row>())
                                {
                                    T obj;
                                    if (rowNum > 0)
                                    {
                                        List<string> fields = new List<string>();
                                        foreach (Cell cell in row.Elements<Cell>())
                                        {
                                            fields.Add(GetCellValue(spreadsheetDocument, cell));
                                        }
                                        obj = (T)constructor.Invoke(new object[] { fields });
                                        ret.Add(obj);
                                    }
                                    rowNum++;
                                }
                            }
                            
                        }else
                        {
                            throw new Exception();
                        }
                        return ret;
                    }else
                    {
                        throw new Exception();
                    }
                }
            }else
            {
                throw new Exception();
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

    }
}
*/