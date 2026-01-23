using System.Runtime.InteropServices;

namespace ExcelAddIn1;

[ComVisible(true)]
public interface IAddInUtilities
{
	void Copy();

	void Cut();

	void Undo();

	void Redo();

	void PasteInsert();

	void PasteNumFormat();

	void PasteLinks();

	void PasteExact();

	void PasteDuplicate();

	void PasteTranspose();

	void CycleNumber();

	void CycleCurrency();

	void CycleForeign();

	void CyclePercent();

	void CycleMultiple();

	void CycleDate();

	void CycleBinary();

	void CycleRatio();

	void AutoColorCycle();

	void NoAutoColor();

	void BlueBlackToggle();

	void CycleFontColor();

	void CycleFill();

	void CycleBorderColor();

	void CycleChartColor();

	void CycleAlternateShading();

	void CenterToggle();

	void HorizAlign();

	void VertAlign();

	void LeftIndent();

	void RightIndent();

	void BorderTheme();

	void BorderTop();

	void BorderBottom();

	void BorderLeft();

	void BorderRight();

	void BorderOutline();

	void BorderInside();

	void BorderNone();

	void FontCycleStyle();

	void FontCycleSize();

	void Underline();

	void CycleCase();

	void CycleList();

	void LeaderDots();

	void SumBar();

	void FootnoteCycle();

	void FootnoteToggle();

	void WrapText();

	void PaintbrushCapture();

	void PaintbrushApply();

	void IncrDecimal();

	void DecrDecimal();

	void ShiftDecimalLeft();

	void ShiftDecimalRight();

	void IncrFont();

	void DecrFont();

	void IncrTableSize();

	void DecrTableSize();

	void AutoColorSelection();

	void AutoColorSheet();

	void AutoColorWorkbook();

	void ProPrecedents();

	void ProDependents();

	void LastAuditedCell();

	void AutoPrecedentsToggle();

	void AutoDependentsToggle();

	void ShowAllPrecedents();

	void ShowAllDependents();

	void ClearArrows();

	void ProCopyRight();

	void ProCopyDown();

	void CheckForErrors();

	void SimplifyFormula();

	void CleanFormula();

	void ConvFormula();

	void Flatten();

	void FlipSign();

	void Untranspose();

	void CommentFormula();

	void WrapParentheses();

	void AutoFillDates();

	void ToggleGridlines();

	void HidePageBreaks();

	void SmartPrintArea();

	void MaximizeWorkspace();

	void ZoomIn();

	void ZoomOut();

	void QuickSave();

	void QuickSaveAll();

	void QuickSaveAs();

	void QuickSaveUp();

	void Reopen();

	void CommentDelete();

	void NextWatch();

	void PrevWatch();

	void AddWatch();

	void RemoveWatch();

	void FirstSheet();

	void LastSheet();

	void NextSheet();

	void PrevSheet();

	void SheetActivate();

	void SheetMoveLeft();

	void SheetMoveRight();

	void GoToMin();

	void GoToMax();

	void CycleRowHeight();

	void CycleColWidth();

	void AutoFitRow();

	void AutoFitCol();

	void InsertRow();

	void InsertCol();

	void DeleteRow();

	void DeleteCol();

	void GroupRow();

	void GroupCol();

	void UngroupRow();

	void UngroupCol();

	void HideRow();

	void HideCol();

	void UnhideRow();

	void UnhideCol();

	void ExpandRows();

	void ExpandCols();

	void CollapseRows();

	void CollapseCols();

	void ProperHide();

	void RowColCopy();

	void RowColPaste();

	void StyleCycle1();

	void StyleCycle2();

	void StyleCycle3();

	void StyleCycle4();

	void StyleCycle5();

	void StyleCycle6();

	void StyleCycle7();

	void StyleCycle8();

	void UniformRange();

	void PasteMatchWidth();

	void PasteMatchHeight();

	void PasteMatchBoth();

	void PasteMatchNone();
}
