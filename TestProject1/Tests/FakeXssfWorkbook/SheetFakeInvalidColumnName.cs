using NPOI.SS.UserModel;
using TestProject1.Tests.FakeXssfWorkbook.BaseFake;

namespace TestProject1.Tests.FakeXssfWorkbook
{
    public class SheetFakeInvalidColumnName : SheetFake, ISheet
    {
        public SheetFakeInvalidColumnName(short lastCellNum) : base(lastCellNum)
        {
        }

        IRow ISheet.GetRow(int rownum)
        {
            return new RowFakeInvalidColumnName(lastCellNum);
        }
    }
}
