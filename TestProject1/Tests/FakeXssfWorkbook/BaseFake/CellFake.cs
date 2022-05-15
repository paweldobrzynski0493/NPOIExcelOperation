using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;

namespace TestProject1.Tests.FakeXssfWorkbook.BaseFake
{
    public class CellFake : ICell
    {
        public string stringCellValue { get; set; }

        public virtual void SetStringCellValue(int cellnum)
        {
            stringCellValue = typeof(TestExcelModel).GetProperties()[cellnum].Name;
        }

        public int ColumnIndex => throw new NotImplementedException();

        public int RowIndex => throw new NotImplementedException();

        public ISheet Sheet => throw new NotImplementedException();

        public IRow Row => throw new NotImplementedException();

        public CellType CellType => throw new NotImplementedException();

        public CellType CachedFormulaResultType => throw new NotImplementedException();

        public string CellFormula { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public double NumericCellValue => throw new NotImplementedException();

        public DateTime DateCellValue => throw new NotImplementedException();

        public IRichTextString RichStringCellValue => throw new NotImplementedException();

        public byte ErrorCellValue => throw new NotImplementedException();

        public string StringCellValue => stringCellValue;

        public bool BooleanCellValue => throw new NotImplementedException();

        public ICellStyle CellStyle { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public CellAddress Address => throw new NotImplementedException();

        public IComment CellComment { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public IHyperlink Hyperlink { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public CellRangeAddress ArrayFormulaRange => throw new NotImplementedException();

        public bool IsPartOfArrayFormulaGroup => throw new NotImplementedException();

        public bool IsMergedCell => throw new NotImplementedException();

        public ICell CopyCellTo(int targetIndex)
        {
            throw new NotImplementedException();
        }

        public CellType GetCachedFormulaResultTypeEnum()
        {
            throw new NotImplementedException();
        }

        public void RemoveCellComment()
        {
            throw new NotImplementedException();
        }

        public void RemoveHyperlink()
        {
            throw new NotImplementedException();
        }

        public void SetAsActiveCell()
        {
            throw new NotImplementedException();
        }

        public void SetBlank()
        {
            throw new NotImplementedException();
        }

        public void SetCellErrorValue(byte value)
        {
            throw new NotImplementedException();
        }

        public void SetCellFormula(string formula)
        {
            throw new NotImplementedException();
        }

        public void SetCellType(CellType cellType)
        {
            throw new NotImplementedException();
        }

        public void SetCellValue(double value)
        {
            throw new NotImplementedException();
        }

        public void SetCellValue(DateTime value)
        {
            throw new NotImplementedException();
        }

        public void SetCellValue(IRichTextString value)
        {
            throw new NotImplementedException();
        }

        public void SetCellValue(string value)
        {
            throw new NotImplementedException();
        }

        public void SetCellValue(bool value)
        {
            throw new NotImplementedException();
        }
    }
}
