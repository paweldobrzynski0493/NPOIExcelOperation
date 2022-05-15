using Microsoft.AspNetCore.Mvc;
using WebApplication4.Services;

namespace WebApplication4.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly IXssfWorkbook _xssfWorkboo;

        public ExcelController(IXssfWorkbook xssfWorkboo)
        {
            _xssfWorkboo = xssfWorkboo;
        }

        [HttpGet(Name = "GetExcelModel")]
        public ConvertionResult<ExcelModel> Get()
        {
            var result = new ConvertionResult<ExcelModel>();

            using (var stream = new FileStream("Excel.xlsx", FileMode.Open))
            {
                var sut = new ExcelOperation(_xssfWorkboo);
                result = sut.ReadExcel<ExcelModel>(stream);
            }

            return result;
        }
    }
}