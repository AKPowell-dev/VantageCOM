using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Audit;
using ExcelAddIn1.Audit.TraceDialogs;
using ExcelAddIn1.Audit.TraceDialogs.Dependents;
using ExcelAddIn1.Audit.TraceDialogs.Precedents;
using ExcelAddIn1.Audit.Visualizations;
using ExcelAddIn1.Charts;
using ExcelAddIn1.Comments;
using ExcelAddIn1.Data;
using ExcelAddIn1.Format;
using ExcelAddIn1.Formulas;
using ExcelAddIn1.Links;
using ExcelAddIn1.Model;
using ExcelAddIn1.RowsColumns;
using ExcelAddIn1.Sheets;
using ExcelAddIn1.UndoRedo;
using ExcelAddIn1.View;
using ExcelAddIn1.Workbook;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.AutoDual)]
public sealed class clsRouter : IAddInUtilities
{
	public const string ClassId = "88903157-450b-41ba-9738-42f8d23d77d2";

	public const string InterfaceId = "47592a98-0756-4a04-8d25-78d86359a2f1";

	public const string EventsId = "68fe24d2-bb2d-430a-87ca-5b0531bb462f";

	public void Copy()
	{
		try
		{
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			if (application.Selection is Range)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				try
				{
					((Range)application.Selection).Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
					Paste.Copy();
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					MessageBox.Show(ex2.Message, VH.A(43304), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					ProjectData.ClearProjectError();
				}
			}
			else if (application.CommandBars.GetEnabledMso(VH.A(224)))
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				application.CommandBars.ExecuteMso(VH.A(224));
			}
			application = null;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			MessageBox.Show(ex4.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
	}

	void IAddInUtilities.Copy()
	{
		//ILSpy generated this explicit interface implementation from .override directive in Copy
		this.Copy();
	}

	public void Cut()
	{
		try
		{
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			if (application.Selection is Range)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				try
				{
					((Range)application.Selection).Cut(RuntimeHelpers.GetObjectValue(Missing.Value));
					Paste.Cut();
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					MessageBox.Show(ex2.Message, VH.A(43304), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					ProjectData.ClearProjectError();
				}
			}
			else if (application.CommandBars.GetEnabledMso(VH.A(197247)))
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				application.CommandBars.ExecuteMso(VH.A(197247));
			}
			application = null;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			MessageBox.Show(ex4.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
	}

	void IAddInUtilities.Cut()
	{
		//ILSpy generated this explicit interface implementation from .override directive in Cut
		this.Cut();
	}

	public void Undo()
	{
		ExcelAddIn1.UndoRedo.Core.Undo();
	}

	void IAddInUtilities.Undo()
	{
		//ILSpy generated this explicit interface implementation from .override directive in Undo
		this.Undo();
	}

	public void Redo()
	{
		ExcelAddIn1.UndoRedo.Core.Redo();
	}

	void IAddInUtilities.Redo()
	{
		//ILSpy generated this explicit interface implementation from .override directive in Redo
		this.Redo();
	}

	public void CycleNumber()
	{
		NumberFormat.CycleNumber();
	}

	void IAddInUtilities.CycleNumber()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleNumber
		this.CycleNumber();
	}

	public void CycleCurrency()
	{
		NumberFormat.CycleCurrency();
	}

	void IAddInUtilities.CycleCurrency()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleCurrency
		this.CycleCurrency();
	}

	public void CycleForeign()
	{
		NumberFormat.CycleForeign();
	}

	void IAddInUtilities.CycleForeign()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleForeign
		this.CycleForeign();
	}

	public void CyclePercent()
	{
		NumberFormat.CyclePercent();
	}

	void IAddInUtilities.CyclePercent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CyclePercent
		this.CyclePercent();
	}

	public void CycleMultiple()
	{
		NumberFormat.CycleMultiple();
	}

	void IAddInUtilities.CycleMultiple()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleMultiple
		this.CycleMultiple();
	}

	public void CycleDate()
	{
		NumberFormat.CycleDate();
	}

	void IAddInUtilities.CycleDate()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleDate
		this.CycleDate();
	}

	public void CycleBinary()
	{
		NumberFormat.CycleBinary();
	}

	void IAddInUtilities.CycleBinary()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleBinary
		this.CycleBinary();
	}

	public void CycleRatio()
	{
		NumberFormat.CycleRatio();
	}

	void IAddInUtilities.CycleRatio()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleRatio
		this.CycleRatio();
	}

	public void AutoColorCycle()
	{
		FontColor.CycleAutoColors();
	}

	void IAddInUtilities.AutoColorCycle()
	{
		//ILSpy generated this explicit interface implementation from .override directive in AutoColorCycle
		this.AutoColorCycle();
	}

	public void NoAutoColor()
	{
	}

	void IAddInUtilities.NoAutoColor()
	{
		//ILSpy generated this explicit interface implementation from .override directive in NoAutoColor
		this.NoAutoColor();
	}

	public void BlueBlackToggle()
	{
		FontColor.BlueBlackToggle();
	}

	void IAddInUtilities.BlueBlackToggle()
	{
		//ILSpy generated this explicit interface implementation from .override directive in BlueBlackToggle
		this.BlueBlackToggle();
	}

	public void CycleFontColor()
	{
		FontColor.Cycle();
	}

	void IAddInUtilities.CycleFontColor()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleFontColor
		this.CycleFontColor();
	}

	public void CycleFill()
	{
		FillColor.Cycle();
	}

	void IAddInUtilities.CycleFill()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleFill
		this.CycleFill();
	}

	public void CycleBorderColor()
	{
		ExcelAddIn1.Format.Borders.CycleColor();
	}

	void IAddInUtilities.CycleBorderColor()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleBorderColor
		this.CycleBorderColor();
	}

	public void CycleChartColor()
	{
		CycleColor.Cycle();
	}

	void IAddInUtilities.CycleChartColor()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleChartColor
		this.CycleChartColor();
	}

	public void CycleAlternateShading()
	{
		AlternateShading.Cycle();
	}

	void IAddInUtilities.CycleAlternateShading()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleAlternateShading
		this.CycleAlternateShading();
	}

	public void CenterToggle()
	{
		Alignment.CycleCenter();
	}

	void IAddInUtilities.CenterToggle()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CenterToggle
		this.CenterToggle();
	}

	public void HorizAlign()
	{
		Alignment.CycleHorizontal();
	}

	void IAddInUtilities.HorizAlign()
	{
		//ILSpy generated this explicit interface implementation from .override directive in HorizAlign
		this.HorizAlign();
	}

	public void VertAlign()
	{
		Alignment.CycleVertical();
	}

	void IAddInUtilities.VertAlign()
	{
		//ILSpy generated this explicit interface implementation from .override directive in VertAlign
		this.VertAlign();
	}

	public void LeftIndent()
	{
		Indent.Left();
	}

	void IAddInUtilities.LeftIndent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in LeftIndent
		this.LeftIndent();
	}

	public void RightIndent()
	{
		Indent.Right();
	}

	void IAddInUtilities.RightIndent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in RightIndent
		this.RightIndent();
	}

	public void FontCycleStyle()
	{
		FontStyle.Cycle();
	}

	void IAddInUtilities.FontCycleStyle()
	{
		//ILSpy generated this explicit interface implementation from .override directive in FontCycleStyle
		this.FontCycleStyle();
	}

	public void FontCycleSize()
	{
		FontSize.Cycle();
	}

	void IAddInUtilities.FontCycleSize()
	{
		//ILSpy generated this explicit interface implementation from .override directive in FontCycleSize
		this.FontCycleSize();
	}

	public void BorderTop()
	{
		ExcelAddIn1.Format.Borders.CycleTop();
	}

	void IAddInUtilities.BorderTop()
	{
		//ILSpy generated this explicit interface implementation from .override directive in BorderTop
		this.BorderTop();
	}

	public void BorderBottom()
	{
		ExcelAddIn1.Format.Borders.CycleBottom();
	}

	void IAddInUtilities.BorderBottom()
	{
		//ILSpy generated this explicit interface implementation from .override directive in BorderBottom
		this.BorderBottom();
	}

	public void BorderLeft()
	{
		ExcelAddIn1.Format.Borders.CycleLeft();
	}

	void IAddInUtilities.BorderLeft()
	{
		//ILSpy generated this explicit interface implementation from .override directive in BorderLeft
		this.BorderLeft();
	}

	public void BorderRight()
	{
		ExcelAddIn1.Format.Borders.CycleRight();
	}

	void IAddInUtilities.BorderRight()
	{
		//ILSpy generated this explicit interface implementation from .override directive in BorderRight
		this.BorderRight();
	}

	public void BorderOutline()
	{
		ExcelAddIn1.Format.Borders.Outside();
	}

	void IAddInUtilities.BorderOutline()
	{
		//ILSpy generated this explicit interface implementation from .override directive in BorderOutline
		this.BorderOutline();
	}

	public void BorderInside()
	{
		ExcelAddIn1.Format.Borders.Inside();
	}

	void IAddInUtilities.BorderInside()
	{
		//ILSpy generated this explicit interface implementation from .override directive in BorderInside
		this.BorderInside();
	}

	public void BorderNone()
	{
		ExcelAddIn1.Format.Borders.None();
	}

	void IAddInUtilities.BorderNone()
	{
		//ILSpy generated this explicit interface implementation from .override directive in BorderNone
		this.BorderNone();
	}

	public void BorderTheme()
	{
		MessageBox.Show(VH.A(197254), VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
	}

	void IAddInUtilities.BorderTheme()
	{
		//ILSpy generated this explicit interface implementation from .override directive in BorderTheme
		this.BorderTheme();
	}

	public void Underline()
	{
		ExcelAddIn1.Format.Underline.Cycle();
	}

	void IAddInUtilities.Underline()
	{
		//ILSpy generated this explicit interface implementation from .override directive in Underline
		this.Underline();
	}

	public void CycleCase()
	{
		Cases.Cycle();
	}

	void IAddInUtilities.CycleCase()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleCase
		this.CycleCase();
	}

	public void CycleList()
	{
		Lists.Cycle();
	}

	void IAddInUtilities.CycleList()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleList
		this.CycleList();
	}

	public void LeaderDots()
	{
		ExcelAddIn1.Format.Miscellaneous.LeaderDots();
	}

	void IAddInUtilities.LeaderDots()
	{
		//ILSpy generated this explicit interface implementation from .override directive in LeaderDots
		this.LeaderDots();
	}

	public void SumBar()
	{
		ExcelAddIn1.Format.SumBar.Toggle();
	}

	void IAddInUtilities.SumBar()
	{
		//ILSpy generated this explicit interface implementation from .override directive in SumBar
		this.SumBar();
	}

	public void FootnoteCycle()
	{
		Footnotes.Cycle();
	}

	void IAddInUtilities.FootnoteCycle()
	{
		//ILSpy generated this explicit interface implementation from .override directive in FootnoteCycle
		this.FootnoteCycle();
	}

	public void FootnoteToggle()
	{
		Footnotes.Toggle();
	}

	void IAddInUtilities.FootnoteToggle()
	{
		//ILSpy generated this explicit interface implementation from .override directive in FootnoteToggle
		this.FootnoteToggle();
	}

	public void WrapText()
	{
		ExcelAddIn1.Format.Miscellaneous.WrapText();
	}

	void IAddInUtilities.WrapText()
	{
		//ILSpy generated this explicit interface implementation from .override directive in WrapText
		this.WrapText();
	}

	public void PaintbrushCapture()
	{
		Paintbrush.Capture();
	}

	void IAddInUtilities.PaintbrushCapture()
	{
		//ILSpy generated this explicit interface implementation from .override directive in PaintbrushCapture
		this.PaintbrushCapture();
	}

	public void PaintbrushApply()
	{
		Paintbrush.Apply();
	}

	void IAddInUtilities.PaintbrushApply()
	{
		//ILSpy generated this explicit interface implementation from .override directive in PaintbrushApply
		this.PaintbrushApply();
	}

	public void IncrDecimal()
	{
		Decimals.Increase();
	}

	void IAddInUtilities.IncrDecimal()
	{
		//ILSpy generated this explicit interface implementation from .override directive in IncrDecimal
		this.IncrDecimal();
	}

	public void DecrDecimal()
	{
		Decimals.Decrease();
	}

	void IAddInUtilities.DecrDecimal()
	{
		//ILSpy generated this explicit interface implementation from .override directive in DecrDecimal
		this.DecrDecimal();
	}

	public void ShiftDecimalLeft()
	{
		ShiftDecimal.Left();
	}

	void IAddInUtilities.ShiftDecimalLeft()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ShiftDecimalLeft
		this.ShiftDecimalLeft();
	}

	public void ShiftDecimalRight()
	{
		ShiftDecimal.Right();
	}

	void IAddInUtilities.ShiftDecimalRight()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ShiftDecimalRight
		this.ShiftDecimalRight();
	}

	public void IncrFont()
	{
		FontSize.Increase();
	}

	void IAddInUtilities.IncrFont()
	{
		//ILSpy generated this explicit interface implementation from .override directive in IncrFont
		this.IncrFont();
	}

	public void DecrFont()
	{
		FontSize.Decrease();
	}

	void IAddInUtilities.DecrFont()
	{
		//ILSpy generated this explicit interface implementation from .override directive in DecrFont
		this.DecrFont();
	}

	public void IncrTableSize()
	{
		Tables.IncreaseSize();
	}

	void IAddInUtilities.IncrTableSize()
	{
		//ILSpy generated this explicit interface implementation from .override directive in IncrTableSize
		this.IncrTableSize();
	}

	public void DecrTableSize()
	{
		Tables.DecreaseSize();
	}

	void IAddInUtilities.DecrTableSize()
	{
		//ILSpy generated this explicit interface implementation from .override directive in DecrTableSize
		this.DecrTableSize();
	}

	public void AutoColorSelection()
	{
		AutoColor.Selection();
	}

	void IAddInUtilities.AutoColorSelection()
	{
		//ILSpy generated this explicit interface implementation from .override directive in AutoColorSelection
		this.AutoColorSelection();
	}

	public void AutoColorSheet()
	{
		AutoColor.Worksheet();
	}

	void IAddInUtilities.AutoColorSheet()
	{
		//ILSpy generated this explicit interface implementation from .override directive in AutoColorSheet
		this.AutoColorSheet();
	}

	public void AutoColorWorkbook()
	{
		AutoColor.Workbook();
	}

	void IAddInUtilities.AutoColorWorkbook()
	{
		//ILSpy generated this explicit interface implementation from .override directive in AutoColorWorkbook
		this.AutoColorWorkbook();
	}

	public void ProPrecedents()
	{
		ExcelAddIn1.Audit.TraceDialogs.Precedents.Dialog.Show();
	}

	void IAddInUtilities.ProPrecedents()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ProPrecedents
		this.ProPrecedents();
	}

	public void ProDependents()
	{
		ExcelAddIn1.Audit.TraceDialogs.Dependents.Dialog.Show();
	}

	void IAddInUtilities.ProDependents()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ProDependents
		this.ProDependents();
	}

	public void LastAuditedCell()
	{
		ExcelAddIn1.Audit.TraceDialogs.Base.GoToLastAuditedCell();
	}

	void IAddInUtilities.LastAuditedCell()
	{
		//ILSpy generated this explicit interface implementation from .override directive in LastAuditedCell
		this.LastAuditedCell();
	}

	public void AutoPrecedentsToggle()
	{
		AutoTrace.TogglePrecedents();
	}

	void IAddInUtilities.AutoPrecedentsToggle()
	{
		//ILSpy generated this explicit interface implementation from .override directive in AutoPrecedentsToggle
		this.AutoPrecedentsToggle();
	}

	public void AutoDependentsToggle()
	{
		AutoTrace.ToggleDependents();
	}

	void IAddInUtilities.AutoDependentsToggle()
	{
		//ILSpy generated this explicit interface implementation from .override directive in AutoDependentsToggle
		this.AutoDependentsToggle();
	}

	public void ShowAllPrecedents()
	{
		TraceAll.Precedents();
	}

	void IAddInUtilities.ShowAllPrecedents()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ShowAllPrecedents
		this.ShowAllPrecedents();
	}

	public void ShowAllDependents()
	{
		TraceAll.Dependents();
	}

	void IAddInUtilities.ShowAllDependents()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ShowAllDependents
		this.ShowAllDependents();
	}

	public void ClearArrows()
	{
		Arrows.Clear();
	}

	void IAddInUtilities.ClearArrows()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ClearArrows
		this.ClearArrows();
	}

	public void ProCopyRight()
	{
		FastFill.Right();
	}

	void IAddInUtilities.ProCopyRight()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ProCopyRight
		this.ProCopyRight();
	}

	public void ProCopyDown()
	{
		FastFill.Down();
	}

	void IAddInUtilities.ProCopyDown()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ProCopyDown
		this.ProCopyDown();
	}

	public void CheckForErrors()
	{
		ErrorWrap.Toggle();
	}

	void IAddInUtilities.CheckForErrors()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CheckForErrors
		this.CheckForErrors();
	}

	public void SimplifyFormula()
	{
		Evaluate.SimplifyFormula();
	}

	void IAddInUtilities.SimplifyFormula()
	{
		//ILSpy generated this explicit interface implementation from .override directive in SimplifyFormula
		this.SimplifyFormula();
	}

	public void CleanFormula()
	{
		Clean.Selection();
	}

	void IAddInUtilities.CleanFormula()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CleanFormula
		this.CleanFormula();
	}

	public void ConvFormula()
	{
		Anchor.Cycle();
	}

	void IAddInUtilities.ConvFormula()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ConvFormula
		this.ConvFormula();
	}

	public void Flatten()
	{
		ExcelAddIn1.Formulas.Flatten.Selection();
	}

	void IAddInUtilities.Flatten()
	{
		//ILSpy generated this explicit interface implementation from .override directive in Flatten
		this.Flatten();
	}

	public void FlipSign()
	{
		ExcelAddIn1.Formulas.FlipSign.Go();
	}

	void IAddInUtilities.FlipSign()
	{
		//ILSpy generated this explicit interface implementation from .override directive in FlipSign
		this.FlipSign();
	}

	public void Untranspose()
	{
		ExcelAddIn1.Formulas.Untranspose.Go();
	}

	void IAddInUtilities.Untranspose()
	{
		//ILSpy generated this explicit interface implementation from .override directive in Untranspose
		this.Untranspose();
	}

	public void CommentFormula()
	{
		ExcelAddIn1.Formulas.Comment.Cells();
	}

	void IAddInUtilities.CommentFormula()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CommentFormula
		this.CommentFormula();
	}

	public void WrapParentheses()
	{
		ExcelAddIn1.Formulas.WrapParentheses.Go();
	}

	void IAddInUtilities.WrapParentheses()
	{
		//ILSpy generated this explicit interface implementation from .override directive in WrapParentheses
		this.WrapParentheses();
	}

	public void AutoFillDates()
	{
		AutoFill.Dates();
	}

	void IAddInUtilities.AutoFillDates()
	{
		//ILSpy generated this explicit interface implementation from .override directive in AutoFillDates
		this.AutoFillDates();
	}

	public void PasteInsert()
	{
		Paste.Insert();
	}

	void IAddInUtilities.PasteInsert()
	{
		//ILSpy generated this explicit interface implementation from .override directive in PasteInsert
		this.PasteInsert();
	}

	public void PasteNumFormat()
	{
		Paste.NumberFormats();
	}

	void IAddInUtilities.PasteNumFormat()
	{
		//ILSpy generated this explicit interface implementation from .override directive in PasteNumFormat
		this.PasteNumFormat();
	}

	public void PasteLinks()
	{
		Paste.Links();
	}

	void IAddInUtilities.PasteLinks()
	{
		//ILSpy generated this explicit interface implementation from .override directive in PasteLinks
		this.PasteLinks();
	}

	public void PasteExact()
	{
		Paste.Exact(trans: false);
	}

	void IAddInUtilities.PasteExact()
	{
		//ILSpy generated this explicit interface implementation from .override directive in PasteExact
		this.PasteExact();
	}

	public void PasteDuplicate()
	{
		Paste.Duplicate();
	}

	void IAddInUtilities.PasteDuplicate()
	{
		//ILSpy generated this explicit interface implementation from .override directive in PasteDuplicate
		this.PasteDuplicate();
	}

	public void PasteTranspose()
	{
		Paste.Transpose();
	}

	void IAddInUtilities.PasteTranspose()
	{
		//ILSpy generated this explicit interface implementation from .override directive in PasteTranspose
		this.PasteTranspose();
	}

	public void ToggleGridlines()
	{
		ExcelAddIn1.View.Gridlines.Toggle();
	}

	void IAddInUtilities.ToggleGridlines()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ToggleGridlines
		this.ToggleGridlines();
	}

	public void HidePageBreaks()
	{
		PrintAreas.HidePageBreaks();
	}

	void IAddInUtilities.HidePageBreaks()
	{
		//ILSpy generated this explicit interface implementation from .override directive in HidePageBreaks
		this.HidePageBreaks();
	}

	public void SmartPrintArea()
	{
		PrintAreas.SmartPrintArea();
	}

	void IAddInUtilities.SmartPrintArea()
	{
		//ILSpy generated this explicit interface implementation from .override directive in SmartPrintArea
		this.SmartPrintArea();
	}

	public void MaximizeWorkspace()
	{
		Workspace.Maximize(blnRequireAuthentication: true);
	}

	void IAddInUtilities.MaximizeWorkspace()
	{
		//ILSpy generated this explicit interface implementation from .override directive in MaximizeWorkspace
		this.MaximizeWorkspace();
	}

	public void ZoomIn()
	{
		Zoom.ZoomIn();
	}

	void IAddInUtilities.ZoomIn()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ZoomIn
		this.ZoomIn();
	}

	public void ZoomOut()
	{
		Zoom.ZoomOut();
	}

	void IAddInUtilities.ZoomOut()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ZoomOut
		this.ZoomOut();
	}

	public void QuickSave()
	{
		ExcelAddIn1.Workbook.QuickSave.Save();
	}

	void IAddInUtilities.QuickSave()
	{
		//ILSpy generated this explicit interface implementation from .override directive in QuickSave
		this.QuickSave();
	}

	public void QuickSaveAll()
	{
		ExcelAddIn1.Workbook.QuickSave.SaveAll();
	}

	void IAddInUtilities.QuickSaveAll()
	{
		//ILSpy generated this explicit interface implementation from .override directive in QuickSaveAll
		this.QuickSaveAll();
	}

	public void QuickSaveAs()
	{
		ExcelAddIn1.Workbook.QuickSave.SaveAs();
	}

	void IAddInUtilities.QuickSaveAs()
	{
		//ILSpy generated this explicit interface implementation from .override directive in QuickSaveAs
		this.QuickSaveAs();
	}

	public void QuickSaveUp()
	{
		ExcelAddIn1.Workbook.QuickSave.SaveUp();
	}

	void IAddInUtilities.QuickSaveUp()
	{
		//ILSpy generated this explicit interface implementation from .override directive in QuickSaveUp
		this.QuickSaveUp();
	}

	public void Reopen()
	{
		ExcelAddIn1.Workbook.Miscellaneous.Reopen();
	}

	void IAddInUtilities.Reopen()
	{
		//ILSpy generated this explicit interface implementation from .override directive in Reopen
		this.Reopen();
	}

	public void CommentDelete()
	{
		CleanUp.Delete();
	}

	void IAddInUtilities.CommentDelete()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CommentDelete
		this.CommentDelete();
	}

	public void AddWatch()
	{
		ExcelAddIn1.Audit.Watches.Add();
	}

	void IAddInUtilities.AddWatch()
	{
		//ILSpy generated this explicit interface implementation from .override directive in AddWatch
		this.AddWatch();
	}

	public void NextWatch()
	{
		ExcelAddIn1.Audit.Watches.GoToNext();
	}

	void IAddInUtilities.NextWatch()
	{
		//ILSpy generated this explicit interface implementation from .override directive in NextWatch
		this.NextWatch();
	}

	public void PrevWatch()
	{
		ExcelAddIn1.Audit.Watches.GoToPrevious();
	}

	void IAddInUtilities.PrevWatch()
	{
		//ILSpy generated this explicit interface implementation from .override directive in PrevWatch
		this.PrevWatch();
	}

	public void RemoveWatch()
	{
		ExcelAddIn1.Audit.Watches.Remove();
	}

	void IAddInUtilities.RemoveWatch()
	{
		//ILSpy generated this explicit interface implementation from .override directive in RemoveWatch
		this.RemoveWatch();
	}

	public void FirstSheet()
	{
		Navigate.A();
	}

	void IAddInUtilities.FirstSheet()
	{
		//ILSpy generated this explicit interface implementation from .override directive in FirstSheet
		this.FirstSheet();
	}

	public void LastSheet()
	{
		Navigate.B();
	}

	void IAddInUtilities.LastSheet()
	{
		//ILSpy generated this explicit interface implementation from .override directive in LastSheet
		this.LastSheet();
	}

	public void NextSheet()
	{
		Navigate.C();
	}

	void IAddInUtilities.NextSheet()
	{
		//ILSpy generated this explicit interface implementation from .override directive in NextSheet
		this.NextSheet();
	}

	public void PrevSheet()
	{
		Navigate.D();
	}

	void IAddInUtilities.PrevSheet()
	{
		//ILSpy generated this explicit interface implementation from .override directive in PrevSheet
		this.PrevSheet();
	}

	public void SheetActivate()
	{
		Navigate.E();
	}

	void IAddInUtilities.SheetActivate()
	{
		//ILSpy generated this explicit interface implementation from .override directive in SheetActivate
		this.SheetActivate();
	}

	public void SheetMoveLeft()
	{
		Move.A();
	}

	void IAddInUtilities.SheetMoveLeft()
	{
		//ILSpy generated this explicit interface implementation from .override directive in SheetMoveLeft
		this.SheetMoveLeft();
	}

	public void SheetMoveRight()
	{
		Move.B();
	}

	void IAddInUtilities.SheetMoveRight()
	{
		//ILSpy generated this explicit interface implementation from .override directive in SheetMoveRight
		this.SheetMoveRight();
	}

	public void GoToMin()
	{
		MinMax.A();
	}

	void IAddInUtilities.GoToMin()
	{
		//ILSpy generated this explicit interface implementation from .override directive in GoToMin
		this.GoToMin();
	}

	public void GoToMax()
	{
		MinMax.B();
	}

	void IAddInUtilities.GoToMax()
	{
		//ILSpy generated this explicit interface implementation from .override directive in GoToMax
		this.GoToMax();
	}

	public void CycleRowHeight()
	{
		CellSize.CycleRowHeight();
	}

	void IAddInUtilities.CycleRowHeight()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleRowHeight
		this.CycleRowHeight();
	}

	public void CycleColWidth()
	{
		CellSize.CycleColumnWidth();
	}

	void IAddInUtilities.CycleColWidth()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleColWidth
		this.CycleColWidth();
	}

	public void AutoFitRow()
	{
		AutoFit.Height();
	}

	void IAddInUtilities.AutoFitRow()
	{
		//ILSpy generated this explicit interface implementation from .override directive in AutoFitRow
		this.AutoFitRow();
	}

	public void AutoFitCol()
	{
		AutoFit.Width();
	}

	void IAddInUtilities.AutoFitCol()
	{
		//ILSpy generated this explicit interface implementation from .override directive in AutoFitCol
		this.AutoFitCol();
	}

	public void InsertRow()
	{
		Insert.Row();
	}

	void IAddInUtilities.InsertRow()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InsertRow
		this.InsertRow();
	}

	public void InsertCol()
	{
		Insert.Column();
	}

	void IAddInUtilities.InsertCol()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InsertCol
		this.InsertCol();
	}

	public void DeleteRow()
	{
		Delete.Row();
	}

	void IAddInUtilities.DeleteRow()
	{
		//ILSpy generated this explicit interface implementation from .override directive in DeleteRow
		this.DeleteRow();
	}

	public void DeleteCol()
	{
		Delete.Column();
	}

	void IAddInUtilities.DeleteCol()
	{
		//ILSpy generated this explicit interface implementation from .override directive in DeleteCol
		this.DeleteCol();
	}

	public void GroupRow()
	{
		Group.Rows();
	}

	void IAddInUtilities.GroupRow()
	{
		//ILSpy generated this explicit interface implementation from .override directive in GroupRow
		this.GroupRow();
	}

	public void GroupCol()
	{
		Group.Columns();
	}

	void IAddInUtilities.GroupCol()
	{
		//ILSpy generated this explicit interface implementation from .override directive in GroupCol
		this.GroupCol();
	}

	public void UngroupRow()
	{
		Ungroup.Rows();
	}

	void IAddInUtilities.UngroupRow()
	{
		//ILSpy generated this explicit interface implementation from .override directive in UngroupRow
		this.UngroupRow();
	}

	public void UngroupCol()
	{
		Ungroup.Columns();
	}

	void IAddInUtilities.UngroupCol()
	{
		//ILSpy generated this explicit interface implementation from .override directive in UngroupCol
		this.UngroupCol();
	}

	public void HideRow()
	{
		Hide.Rows();
	}

	void IAddInUtilities.HideRow()
	{
		//ILSpy generated this explicit interface implementation from .override directive in HideRow
		this.HideRow();
	}

	public void HideCol()
	{
		Hide.Columns();
	}

	void IAddInUtilities.HideCol()
	{
		//ILSpy generated this explicit interface implementation from .override directive in HideCol
		this.HideCol();
	}

	public void UnhideRow()
	{
		Unhide.Rows();
	}

	void IAddInUtilities.UnhideRow()
	{
		//ILSpy generated this explicit interface implementation from .override directive in UnhideRow
		this.UnhideRow();
	}

	public void UnhideCol()
	{
		Unhide.Columns();
	}

	void IAddInUtilities.UnhideCol()
	{
		//ILSpy generated this explicit interface implementation from .override directive in UnhideCol
		this.UnhideCol();
	}

	public void ExpandRows()
	{
		Expand.Rows();
	}

	void IAddInUtilities.ExpandRows()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ExpandRows
		this.ExpandRows();
	}

	public void ExpandCols()
	{
		Expand.Columns();
	}

	void IAddInUtilities.ExpandCols()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ExpandCols
		this.ExpandCols();
	}

	public void CollapseRows()
	{
		Collapse.Rows();
	}

	void IAddInUtilities.CollapseRows()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CollapseRows
		this.CollapseRows();
	}

	public void CollapseCols()
	{
		Collapse.Columns();
	}

	void IAddInUtilities.CollapseCols()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CollapseCols
		this.CollapseCols();
	}

	public void ProperHide()
	{
		Hide.ProperHide();
	}

	void IAddInUtilities.ProperHide()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ProperHide
		this.ProperHide();
	}

	public void RowColCopy()
	{
		Info.Copy();
	}

	void IAddInUtilities.RowColCopy()
	{
		//ILSpy generated this explicit interface implementation from .override directive in RowColCopy
		this.RowColCopy();
	}

	public void RowColPaste()
	{
		Info.Paste();
	}

	void IAddInUtilities.RowColPaste()
	{
		//ILSpy generated this explicit interface implementation from .override directive in RowColPaste
		this.RowColPaste();
	}

	public void StyleCycle1()
	{
		ExcelAddIn1.Format.Styles.CycleCustom1();
	}

	void IAddInUtilities.StyleCycle1()
	{
		//ILSpy generated this explicit interface implementation from .override directive in StyleCycle1
		this.StyleCycle1();
	}

	public void StyleCycle2()
	{
		ExcelAddIn1.Format.Styles.CycleCustom2();
	}

	void IAddInUtilities.StyleCycle2()
	{
		//ILSpy generated this explicit interface implementation from .override directive in StyleCycle2
		this.StyleCycle2();
	}

	public void StyleCycle3()
	{
		ExcelAddIn1.Format.Styles.CycleCustom3();
	}

	void IAddInUtilities.StyleCycle3()
	{
		//ILSpy generated this explicit interface implementation from .override directive in StyleCycle3
		this.StyleCycle3();
	}

	public void StyleCycle4()
	{
		ExcelAddIn1.Format.Styles.CycleCustom4();
	}

	void IAddInUtilities.StyleCycle4()
	{
		//ILSpy generated this explicit interface implementation from .override directive in StyleCycle4
		this.StyleCycle4();
	}

	public void StyleCycle5()
	{
		ExcelAddIn1.Format.Styles.CycleCustom5();
	}

	void IAddInUtilities.StyleCycle5()
	{
		//ILSpy generated this explicit interface implementation from .override directive in StyleCycle5
		this.StyleCycle5();
	}

	public void StyleCycle6()
	{
		ExcelAddIn1.Format.Styles.CycleCustom6();
	}

	void IAddInUtilities.StyleCycle6()
	{
		//ILSpy generated this explicit interface implementation from .override directive in StyleCycle6
		this.StyleCycle6();
	}

	public void StyleCycle7()
	{
		ExcelAddIn1.Format.Styles.CycleCustom7();
	}

	void IAddInUtilities.StyleCycle7()
	{
		//ILSpy generated this explicit interface implementation from .override directive in StyleCycle7
		this.StyleCycle7();
	}

	public void StyleCycle8()
	{
		ExcelAddIn1.Format.Styles.CycleCustom8();
	}

	void IAddInUtilities.StyleCycle8()
	{
		//ILSpy generated this explicit interface implementation from .override directive in StyleCycle8
		this.StyleCycle8();
	}

	public void UniformRange()
	{
		Uniformulas.Apply();
	}

	void IAddInUtilities.UniformRange()
	{
		//ILSpy generated this explicit interface implementation from .override directive in UniformRange
		this.UniformRange();
	}

	public void PasteMatchWidth()
	{
		Export.A();
	}

	void IAddInUtilities.PasteMatchWidth()
	{
		//ILSpy generated this explicit interface implementation from .override directive in PasteMatchWidth
		this.PasteMatchWidth();
	}

	public void PasteMatchNone()
	{
		Export.B();
	}

	void IAddInUtilities.PasteMatchNone()
	{
		//ILSpy generated this explicit interface implementation from .override directive in PasteMatchNone
		this.PasteMatchNone();
	}

	public void PasteMatchHeight()
	{
		throw new NotImplementedException();
	}

	void IAddInUtilities.PasteMatchHeight()
	{
		//ILSpy generated this explicit interface implementation from .override directive in PasteMatchHeight
		this.PasteMatchHeight();
	}

	public void PasteMatchBoth()
	{
		throw new NotImplementedException();
	}

	void IAddInUtilities.PasteMatchBoth()
	{
		//ILSpy generated this explicit interface implementation from .override directive in PasteMatchBoth
		this.PasteMatchBoth();
	}
}
