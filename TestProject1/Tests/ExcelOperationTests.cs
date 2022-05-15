using NPOI.SS.UserModel;
using NSubstitute;
using NUnit.Framework;
using FluentAssertions;
using WebApplication4.Services;
using System.IO;
using TestProject1.Tests.FakeXssfWorkbook;
using TestProject1.Tests.FakeXssfWorkbook.BaseFake;

namespace TestProject1.Tests
{
    [TestFixture]
    public class ExcelOperationTests
    {
        [Test]
        public void ExcelOperation_ReadExcel_WhenNumerOfColumnsIsNotEqualToNumberOfProperties_ShouldReturnExpectedErrorMessage()
        {
            //setup
            var cellNumerbDifferent = (short)(typeof(TestExcelModel).GetProperties().Length + 1);
            var stream = new MemoryStream();
            var xssfWorkbook = Substitute.For<IXssfWorkbook>();
            var workbook = Substitute.For<IWorkbook>();
            workbook.GetSheetAt(Arg.Any<int>()).Returns(new SheetFake(cellNumerbDifferent));
            xssfWorkbook.CreateXSSFWorkbook(stream).Returns(workbook);

            var sut = new ExcelOperation(xssfWorkbook);

            //execute
            var result = sut.ReadExcel<TestExcelModel>(stream);

            //check
            result.ErrorMessage.Should().Be("Niepoprawna liczba kolumn. Oczekiwano nastêpuj¹cej liczby kolumn: " + (short)(typeof(TestExcelModel).GetProperties().Length));
        }

        [Test]
        public void ExcelOperation_ReadExcel_WhenColumnNamNotExistInPropertiesNames_ShouldReturnExpectedErrorMessage()
        {
            //setup
            var cellNumerbDifferent = (short)(typeof(TestExcelModel).GetProperties().Length);
            var stream = new MemoryStream();
            var xssfWorkbook = Substitute.For<IXssfWorkbook>();
            var workbook = Substitute.For<IWorkbook>();
            workbook.GetSheetAt(Arg.Any<int>()).Returns(new SheetFakeInvalidColumnName(cellNumerbDifferent));
            xssfWorkbook.CreateXSSFWorkbook(stream).Returns(workbook);

            var sut = new ExcelOperation(xssfWorkbook);

            //execute
            var result = sut.ReadExcel<TestExcelModel>(stream);

            //check
            result.ErrorMessage.Should().Be("Niepoprawna nazwa kolumny: InvalidClumnName w komórce nr: 1");
        }
    }
}