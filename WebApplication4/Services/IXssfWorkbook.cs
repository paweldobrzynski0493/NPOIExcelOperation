using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace WebApplication4.Services
{
    public interface IXssfWorkbook
    {
        public IWorkbook CreateXSSFWorkbook(Stream stream);
    }
}
