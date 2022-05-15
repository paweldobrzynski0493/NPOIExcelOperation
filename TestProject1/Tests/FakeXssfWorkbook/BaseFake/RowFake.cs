using NPOI.SS.UserModel;
using System.Collections;
using System.Collections.Generic;

namespace TestProject1.Tests.FakeXssfWorkbook.BaseFake
{
    public class RowFake : IRow
    {
        public RowFake(short lastCellNum)
        {
            LastCellNum = lastCellNum;
        }

        public int RowNum { get => throw new System.NotImplementedException(); set => throw new System.NotImplementedException(); }

        public short FirstCellNum => throw new System.NotImplementedException();

        public short LastCellNum { get; set; }

        public int PhysicalNumberOfCells => throw new System.NotImplementedException();

        public bool ZeroHeight { get => throw new System.NotImplementedException(); set => throw new System.NotImplementedException(); }
        public short Height { get => throw new System.NotImplementedException(); set => throw new System.NotImplementedException(); }
        public float HeightInPoints { get => throw new System.NotImplementedException(); set => throw new System.NotImplementedException(); }

        public bool IsFormatted => throw new System.NotImplementedException();

        public ISheet Sheet => throw new System.NotImplementedException();

        public ICellStyle RowStyle { get => throw new System.NotImplementedException(); set => throw new System.NotImplementedException(); }

        public List<ICell> Cells => throw new System.NotImplementedException();

        public int OutlineLevel => throw new System.NotImplementedException();

        public bool? Hidden { get => throw new System.NotImplementedException(); set => throw new System.NotImplementedException(); }
        public bool? Collapsed { get => throw new System.NotImplementedException(); set => throw new System.NotImplementedException(); }

        public virtual ICell GetCell(int cellnum)
        {
            var cellFake = new CellFake();
            cellFake.SetStringCellValue(cellnum);

            return cellFake;
        }

        public ICell CopyCell(int sourceIndex, int targetIndex)
        {
            throw new System.NotImplementedException();
        }

        public IRow CopyRowTo(int targetIndex)
        {
            throw new System.NotImplementedException();
        }

        public ICell CreateCell(int column)
        {
            throw new System.NotImplementedException();
        }

        public ICell CreateCell(int column, CellType type)
        {
            throw new System.NotImplementedException();
        }

        public ICell GetCell(int cellnum, MissingCellPolicy policy)
        {
            throw new System.NotImplementedException();
        }

        public IEnumerator<ICell> GetEnumerator()
        {
            throw new System.NotImplementedException();
        }

        public bool HasCustomHeight()
        {
            throw new System.NotImplementedException();
        }

        public void MoveCell(ICell cell, int newColumn)
        {
            throw new System.NotImplementedException();
        }

        public void RemoveCell(ICell cell)
        {
            throw new System.NotImplementedException();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new System.NotImplementedException();
        }
    }
}
