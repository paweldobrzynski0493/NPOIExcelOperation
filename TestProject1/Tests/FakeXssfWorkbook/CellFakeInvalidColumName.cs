using TestProject1.Tests.FakeXssfWorkbook.BaseFake;

namespace TestProject1.Tests.FakeXssfWorkbook
{
    public class CellFakeInvalidColumName : CellFake
    {
        public override void SetStringCellValue(int cellnum)
        {
            stringCellValue = "InvalidClumnName";
        }
    }
}
