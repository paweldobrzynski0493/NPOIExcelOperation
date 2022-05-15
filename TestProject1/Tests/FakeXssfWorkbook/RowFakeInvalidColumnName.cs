using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestProject1.Tests.FakeXssfWorkbook.BaseFake;

namespace TestProject1.Tests.FakeXssfWorkbook
{
    public class RowFakeInvalidColumnName : RowFake
    {
        public RowFakeInvalidColumnName(short lastCellNum) : base(lastCellNum)
        {
        }

        public override ICell GetCell(int cellnum)
        {
            var cellFake = new CellFakeInvalidColumName();
            cellFake.SetStringCellValue(cellnum);

            return cellFake;
        }
    }
}
