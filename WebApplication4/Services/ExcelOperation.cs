using NPOI.SS.UserModel;
using System.Reflection;

namespace WebApplication4.Services
{
    public class ExcelOperation
    {
        private string ErrorMessage { get; set; }
        private PropertyInfo[] Props { get; set; }

        private readonly IXssfWorkbook _xssfWorkbook;

        public ExcelOperation(IXssfWorkbook xssfWorkbook)
        {
            _xssfWorkbook = xssfWorkbook;
        }

        public ConvertionResult<T> ReadExcel<T>(Stream stream) where T : new()
        {
            var result = new ConvertionResult<T>();
            stream.Position = 0;
            var xssWorkbook = _xssfWorkbook.CreateXSSFWorkbook(stream);
            var sheet = xssWorkbook.GetSheetAt(0);
            var headerRow = sheet.GetRow(0);
            Props = typeof(T).GetProperties();

            if (!IsValidColumnNames(headerRow))
            {
                result.ErrorMessage = ErrorMessage;
                return result;
            }

            for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                var rowResult = new T();

                if (row == null || row.Cells.All(d => d.CellType == CellType.Blank))
                {
                    result.ErrorMessage = "Odczyt nie powiódł się. Znaleziono pusty wiersz nr: " + ++i;
                    return result;
                }

                for (int j = row.FirstCellNum; j < headerRow.LastCellNum; j++)
                {
                    if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) && !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
                    {
                        var cellType = row.GetCell(j).CellType;
                        if (cellType == CellType.Boolean)
                        {
                            Props[j].SetValue(rowResult, row.GetCell(j).BooleanCellValue);
                        }
                        else if (cellType == CellType.String)
                        {
                            Props[j].SetValue(rowResult, row.GetCell(j).StringCellValue);
                        }
                        else if (cellType == CellType.Numeric)
                        {
                            if (Props[j].PropertyType == typeof(bool))
                            {
                                bool.TryParse((row.GetCell(j).ToString()), out bool res);
                                Props[j].SetValue(rowResult, res);
                            }
                            else
                            {
                                //if (Props[j].isn == typeof(string)) //sprawdzić czy pole jest nullowalne w kodzie i jeśli nie jest a przychodzi puste to przerwać!
                                //{

                                //}
                                var t = Nullable.GetUnderlyingType(Props[j].PropertyType) ?? Props[j].PropertyType;
                                Props[j].SetValue(rowResult, (row.GetCell(j).ToString() == null) ? null : Convert.ChangeType(row.GetCell(j).ToString(), t));
                            }
                        }
                        else
                        {
                            var cell = cellType;
                        }
                    }
                }

                result.OutputList.Add(rowResult);
            }

            return result;
        }

        private bool IsValidColumnNames(IRow headerRow) //TODO stringBuilder
        {
            if (headerRow.LastCellNum != Props.Length)
            {
                ErrorMessage = "Niepoprawna liczba kolumn. Oczekiwano następującej liczby kolumn: " + Props.Length;
                return false;
            }

            for (int i = 0; i < headerRow.LastCellNum; i++)
            {
                ICell cell = headerRow.GetCell(i);
                if (cell == null || string.IsNullOrWhiteSpace(cell.ToString()))
                {
                    ErrorMessage = "Brak oczekiwanej nazwy kolumny: " + Props[i].Name + " w komórce nr: " + ++i;
                    return false;
                }

                if (Props[i].Name != cell.StringCellValue)
                {
                    ErrorMessage = "Niepoprawna nazwa kolumny: " + cell.StringCellValue + " w komórce nr: " + ++i;
                    return false;
                }
            }

            return true;
        }
    }
}
