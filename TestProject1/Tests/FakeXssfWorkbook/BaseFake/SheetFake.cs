using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections;
using System.Collections.Generic;

namespace TestProject1.Tests.FakeXssfWorkbook.BaseFake
{
    public class SheetFake : ISheet
    {
        public short lastCellNum { get; set; }

        public SheetFake(short lastCellNum)
        {
            this.lastCellNum = lastCellNum;
        }
        int ISheet.PhysicalNumberOfRows => 2;

        int ISheet.FirstRowNum => 0;

        int ISheet.LastRowNum => 1;

        bool ISheet.ForceFormulaRecalculation { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        int ISheet.DefaultColumnWidth { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        short ISheet.DefaultRowHeight { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        float ISheet.DefaultRowHeightInPoints { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool ISheet.HorizontallyCenter { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool ISheet.VerticallyCenter { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        int ISheet.NumMergedRegions => throw new NotImplementedException();

        List<CellRangeAddress> ISheet.MergedRegions => throw new NotImplementedException();

        bool ISheet.DisplayZeros { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool ISheet.Autobreaks { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool ISheet.DisplayGuts { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool ISheet.FitToPage { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool ISheet.RowSumsBelow { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool ISheet.RowSumsRight { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool ISheet.IsPrintGridlines { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool ISheet.IsPrintRowAndColumnHeadings { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        IPrintSetup ISheet.PrintSetup => throw new NotImplementedException();

        IHeader ISheet.Header => throw new NotImplementedException();

        IFooter ISheet.Footer => throw new NotImplementedException();

        bool ISheet.Protect => throw new NotImplementedException();

        bool ISheet.ScenarioProtect => throw new NotImplementedException();

        short ISheet.TabColorIndex { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        IDrawing ISheet.DrawingPatriarch => throw new NotImplementedException();

        short ISheet.TopRow { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        short ISheet.LeftCol { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        PaneInformation ISheet.PaneInformation => throw new NotImplementedException();

        bool ISheet.DisplayGridlines { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool ISheet.DisplayFormulas { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool ISheet.DisplayRowColHeadings { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool ISheet.IsActive { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        int[] ISheet.RowBreaks => throw new NotImplementedException();

        int[] ISheet.ColumnBreaks => throw new NotImplementedException();

        IWorkbook ISheet.Workbook => throw new NotImplementedException();

        string ISheet.SheetName => throw new NotImplementedException();

        bool ISheet.IsSelected { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        ISheetConditionalFormatting ISheet.SheetConditionalFormatting => throw new NotImplementedException();

        bool ISheet.IsRightToLeft { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        CellRangeAddress ISheet.RepeatingRows { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        CellRangeAddress ISheet.RepeatingColumns { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        CellAddress ISheet.ActiveCell { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        IRow ISheet.GetRow(int rownum)
        {
            return new RowFake(lastCellNum);
        }

        public IEnumerator GetEnumerator()
        {
            throw new NotImplementedException();
        }

        public IEnumerator GetRowEnumerator()
        {
            throw new NotImplementedException();
        }

        int ISheet.AddMergedRegion(CellRangeAddress region)
        {
            throw new NotImplementedException();
        }

        int ISheet.AddMergedRegionUnsafe(CellRangeAddress region)
        {
            throw new NotImplementedException();
        }

        void ISheet.AddValidationData(IDataValidation dataValidation)
        {
            throw new NotImplementedException();
        }

        void ISheet.AutoSizeColumn(int column)
        {
            throw new NotImplementedException();
        }

        void ISheet.AutoSizeColumn(int column, bool useMergedCells)
        {
            throw new NotImplementedException();
        }

        IRow ISheet.CopyRow(int sourceIndex, int targetIndex)
        {
            throw new NotImplementedException();
        }

        ISheet ISheet.CopySheet(string Name)
        {
            throw new NotImplementedException();
        }

        ISheet ISheet.CopySheet(string Name, bool copyStyle)
        {
            throw new NotImplementedException();
        }

        void ISheet.CopyTo(IWorkbook dest, string name, bool copyStyle, bool keepFormulas)
        {
            throw new NotImplementedException();
        }

        IDrawing ISheet.CreateDrawingPatriarch()
        {
            throw new NotImplementedException();
        }

        void ISheet.CreateFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow)
        {
            throw new NotImplementedException();
        }

        void ISheet.CreateFreezePane(int colSplit, int rowSplit)
        {
            throw new NotImplementedException();
        }

        IRow ISheet.CreateRow(int rownum)
        {
            throw new NotImplementedException();
        }

        void ISheet.CreateSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, PanePosition activePane)
        {
            throw new NotImplementedException();
        }

        IComment ISheet.GetCellComment(int row, int column)
        {
            throw new NotImplementedException();
        }

        IComment ISheet.GetCellComment(CellAddress ref1)
        {
            throw new NotImplementedException();
        }

        Dictionary<CellAddress, IComment> ISheet.GetCellComments()
        {
            throw new NotImplementedException();
        }

        int ISheet.GetColumnOutlineLevel(int columnIndex)
        {
            throw new NotImplementedException();
        }

        ICellStyle ISheet.GetColumnStyle(int column)
        {
            throw new NotImplementedException();
        }

        int ISheet.GetColumnWidth(int columnIndex)
        {
            throw new NotImplementedException();
        }

        float ISheet.GetColumnWidthInPixels(int columnIndex)
        {
            throw new NotImplementedException();
        }

        IDataValidationHelper ISheet.GetDataValidationHelper()
        {
            throw new NotImplementedException();
        }

        List<IDataValidation> ISheet.GetDataValidations()
        {
            throw new NotImplementedException();
        }

        IHyperlink ISheet.GetHyperlink(int row, int column)
        {
            throw new NotImplementedException();
        }

        IHyperlink ISheet.GetHyperlink(CellAddress addr)
        {
            throw new NotImplementedException();
        }

        List<IHyperlink> ISheet.GetHyperlinkList()
        {
            throw new NotImplementedException();
        }

        double ISheet.GetMargin(MarginType margin)
        {
            throw new NotImplementedException();
        }

        CellRangeAddress ISheet.GetMergedRegion(int index)
        {
            throw new NotImplementedException();
        }

        void ISheet.GroupColumn(int fromColumn, int toColumn)
        {
            throw new NotImplementedException();
        }

        void ISheet.GroupRow(int fromRow, int toRow)
        {
            throw new NotImplementedException();
        }

        bool ISheet.IsColumnBroken(int column)
        {
            throw new NotImplementedException();
        }

        bool ISheet.IsColumnHidden(int columnIndex)
        {
            throw new NotImplementedException();
        }

        bool ISheet.IsDate1904()
        {
            throw new NotImplementedException();
        }

        bool ISheet.IsMergedRegion(CellRangeAddress mergedRegion)
        {
            throw new NotImplementedException();
        }

        bool ISheet.IsRowBroken(int row)
        {
            throw new NotImplementedException();
        }

        void ISheet.ProtectSheet(string password)
        {
            throw new NotImplementedException();
        }

        ICellRange<ICell> ISheet.RemoveArrayFormula(ICell cell)
        {
            throw new NotImplementedException();
        }

        void ISheet.RemoveColumnBreak(int column)
        {
            throw new NotImplementedException();
        }

        void ISheet.RemoveMergedRegion(int index)
        {
            throw new NotImplementedException();
        }

        void ISheet.RemoveMergedRegions(IList<int> indices)
        {
            throw new NotImplementedException();
        }

        void ISheet.RemoveRow(IRow row)
        {
            throw new NotImplementedException();
        }

        void ISheet.RemoveRowBreak(int row)
        {
            throw new NotImplementedException();
        }

        void ISheet.SetActive(bool value)
        {
            throw new NotImplementedException();
        }

        void ISheet.SetActiveCellRange(int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            throw new NotImplementedException();
        }

        void ISheet.SetActiveCellRange(List<CellRangeAddress8Bit> cellranges, int activeRange, int activeRow, int activeColumn)
        {
            throw new NotImplementedException();
        }

        ICellRange<ICell> ISheet.SetArrayFormula(string formula, CellRangeAddress range)
        {
            throw new NotImplementedException();
        }

        IAutoFilter ISheet.SetAutoFilter(CellRangeAddress range)
        {
            throw new NotImplementedException();
        }

        void ISheet.SetColumnBreak(int column)
        {
            throw new NotImplementedException();
        }

        void ISheet.SetColumnGroupCollapsed(int columnNumber, bool collapsed)
        {
            throw new NotImplementedException();
        }

        void ISheet.SetColumnHidden(int columnIndex, bool hidden)
        {
            throw new NotImplementedException();
        }

        void ISheet.SetColumnWidth(int columnIndex, int width)
        {
            throw new NotImplementedException();
        }

        void ISheet.SetDefaultColumnStyle(int column, ICellStyle style)
        {
            throw new NotImplementedException();
        }

        void ISheet.SetMargin(MarginType margin, double size)
        {
            throw new NotImplementedException();
        }

        void ISheet.SetRowBreak(int row)
        {
            throw new NotImplementedException();
        }

        void ISheet.SetRowGroupCollapsed(int row, bool collapse)
        {
            throw new NotImplementedException();
        }

        void ISheet.SetZoom(int numerator, int denominator)
        {
            throw new NotImplementedException();
        }

        void ISheet.SetZoom(int scale)
        {
            throw new NotImplementedException();
        }

        void ISheet.ShiftRows(int startRow, int endRow, int n)
        {
            throw new NotImplementedException();
        }

        void ISheet.ShiftRows(int startRow, int endRow, int n, bool copyRowHeight, bool resetOriginalRowHeight)
        {
            throw new NotImplementedException();
        }

        void ISheet.ShowInPane(int toprow, int leftcol)
        {
            throw new NotImplementedException();
        }

        void ISheet.UngroupColumn(int fromColumn, int toColumn)
        {
            throw new NotImplementedException();
        }

        void ISheet.UngroupRow(int fromRow, int toRow)
        {
            throw new NotImplementedException();
        }

        void ISheet.ValidateMergedRegions()
        {
            throw new NotImplementedException();
        }
    }
}
