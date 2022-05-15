using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace WebApplication4.Services
{
    public class XssfWorkbook : IXssfWorkbook
    {
        public IWorkbook CreateXSSFWorkbook(Stream stream)
        {
            return new XSSFWorkbook(stream);
        }
    }
}
