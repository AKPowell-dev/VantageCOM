using System;
using System.Drawing;
using System.Reflection;
using System.Resources;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Xml;
using A;
using ExcelAddIn1.Audit;
using ExcelAddIn1.Audit.Check.UI;
using ExcelAddIn1.Audit.TraceDialogs;
using ExcelAddIn1.Audit.TraceDialogs.Dependents;
using ExcelAddIn1.Audit.TraceDialogs.Precedents;
using ExcelAddIn1.Audit.Visualizations;
using ExcelAddIn1.Charts;
using ExcelAddIn1.Charts.GrowthArrow;
using ExcelAddIn1.Charts.MoveDataLabels;
using ExcelAddIn1.Comments;
using ExcelAddIn1.Data;
using ExcelAddIn1.ExcelApp;
using ExcelAddIn1.FastFormats.Charts;
using ExcelAddIn1.Format;
using ExcelAddIn1.FormatPainter;
using ExcelAddIn1.Formulas;
using ExcelAddIn1.Keyboard;
using ExcelAddIn1.Library2;
using ExcelAddIn1.Library2.UI;
using ExcelAddIn1.Library2.Versioning;
using ExcelAddIn1.Links;
using ExcelAddIn1.Model;
using ExcelAddIn1.Publishing;
using ExcelAddIn1.Publishing.Share;
using ExcelAddIn1.RowsColumns;
using ExcelAddIn1.Shapes;
using ExcelAddIn1.Sheets;
using ExcelAddIn1.SuperFind2.UI;
using ExcelAddIn1.UndoRedo;
using ExcelAddIn1.View;
using ExcelAddIn1.Workbook;
using ExcelAddIn1.Workbook.Merge;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.Config;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.Feedback;
using MacabacusMacros.ImportExport;
using MacabacusMacros.Libraries.Manage;
using MacabacusMacros.UI;
using MacabacusMacros.UI.FormsExtensions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;
using stdole;

namespace ExcelAddIn1;

[ComVisible(true)]
public sealed class RibbonXL : IRibbonExtensibility
{
	private IRibbonUI m_A;

	[CompilerGenerated]
	private static bool m_A;

	private static bool BlockExecuteMso
	{
		[CompilerGenerated]
		get
		{
			return RibbonXL.m_A;
		}
		[CompilerGenerated]
		set
		{
			RibbonXL.m_A = value;
		}
	}

	public string GetCustomUI(string ribbonID)
	{
		string outerXml = default(string);
		try
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(J.Ribbon);
			outerXml = xmlDocument.OuterXml;
			return outerXml;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		return outerXml;
	}

	string IRibbonExtensibility.GetCustomUI(string ribbonID)
	{
		//ILSpy generated this explicit interface implementation from .override directive in GetCustomUI
		return this.GetCustomUI(ribbonID);
	}

	public void Ribbon_Load(IRibbonUI ribbonUI)
	{
		this.m_A = ribbonUI;
		KH.A = ribbonUI;
		clsRibbon.Ribbon = ribbonUI;
		Licensing.Authenticate();
	}

	public IPictureDisp CallbackLoadImage(string resourceName)
	{
		return LH.A((Bitmap)new ResourceManager(VH.A(1), Assembly.GetExecutingAssembly()).GetObject(resourceName));
	}

	public string MacabacusKeyTip(IRibbonControl control)
	{
		return clsRibbon.MacabacusTabKeyTipExcel();
	}

	public string GetCustomLabel(IRibbonControl control)
	{
		return clsRibbon.GetCustomLabel(control);
	}

	public int GetItemCount(IRibbonControl control)
	{
		return clsRibbon.GetItemCount(control);
	}

	public Bitmap GetItemImage(IRibbonControl control, int index)
	{
		return clsRibbon.GetItemImage(control, index);
	}

	public void GalleryOnAction(IRibbonControl control, string id, int index)
	{
		clsRibbon.ApplyGalleryStyle(control, index);
	}

	public bool CustomMenuVisible(IRibbonControl control)
	{
		return clsRibbon.CustomMenuVisible(control);
	}

	public bool CustomGalleryVisible(IRibbonControl control)
	{
		return clsRibbon.CustomGalleryVisible(control);
	}

	public string GetStyleScreentip(IRibbonControl control, int index)
	{
		return clsRibbon.GetStyleScreentip(control, index);
	}

	public string GetCycleScreentip(IRibbonControl control)
	{
		return clsRibbon.GetCycleScreentip(control);
	}

	public void DoStyleCycle(IRibbonControl control)
	{
		ExcelAddIn1.Format.Styles.CustomCycle(checked(Conversions.ToInteger(control.Tag) + 1));
	}

	public void DoStyle(ref IRibbonControl control)
	{
		ExcelAddIn1.Format.Styles.DoStyle(control);
	}

	public bool UsePaletteMenus(IRibbonControl control)
	{
		return !clsColors.UseRibbonGalleries;
	}

	public bool UsePaletteGalleries(IRibbonControl control)
	{
		return clsColors.UseRibbonGalleries;
	}

	public string FontColorMenu(IRibbonControl control)
	{
		return clsColors.FontColorMenu(clsColors.AllColorRoles);
	}

	public string FillColorMenu(IRibbonControl control)
	{
		return clsColors.FillColorMenu(clsColors.AllColorRoles);
	}

	public string BorderColorMenu(IRibbonControl control)
	{
		return clsColors.BorderColorMenu(clsColors.AllColorRoles);
	}

	public void DoFontColor(IRibbonControl control)
	{
		FontColor.A(control.Tag);
	}

	public void DoFillColor(IRibbonControl control)
	{
		FillColor.A(control.Tag);
	}

	public void DoBorderColor(IRibbonControl control)
	{
		ExcelAddIn1.Format.Borders.A(control.Tag);
	}

	public Bitmap RibbonColorSquare(IRibbonControl control)
	{
		return clsColors.ColorSquare(control.Tag);
	}

	public Bitmap GetColorSquare(IRibbonControl control, int index)
	{
		return clsColors.GetColorSquare(index);
	}

	public string ColorScreenTip(IRibbonControl control, int index)
	{
		return clsColors.ColorScreenTip(index);
	}

	public void FontColorAction(IRibbonControl control, string id, int index)
	{
		FontColor.A(index);
	}

	public void FillColorAction(IRibbonControl control, string id, int index)
	{
		FillColor.A(index);
	}

	public void BorderColorAction(IRibbonControl control, string id, int index)
	{
		ExcelAddIn1.Format.Borders.A(index);
	}

	public int ColorGalleryCount(IRibbonControl control)
	{
		return clsColors.ColorGalleryCount();
	}

	public Bitmap GetGalleryButtonImage(IRibbonControl control)
	{
		return clsRibbon.RecolorColorButton(control.Id);
	}

	public void FontColorButton(IRibbonControl control)
	{
		FontColor.A();
	}

	public void FillColorButton(IRibbonControl control)
	{
		FillColor.A();
	}

	public void BorderColorButton(IRibbonControl control)
	{
		ExcelAddIn1.Format.Borders.A();
	}

	public void Recolor(IRibbonControl control)
	{
	}

	public bool IsNotProtectedView(IRibbonControl control)
	{
		if (!ExcelAddIn1.Workbook.Miscellaneous.IsProtectedView(SuppressMessages: true))
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return SoftDisable.IsEnabled;
				}
			}
		}
		return false;
	}

	public bool VisPublish(IRibbonControl control)
	{
		return true;
	}

	public bool ShowBetaTools(IRibbonControl control)
	{
		return clsRibbon.ShowBetaTools;
	}

	public string GetScreenTip(IRibbonControl control)
	{
		return ExcelAddIn1.Keyboard.Shortcuts.HotkeyScreenTip(control.Id);
	}

	public string GetSuperTip(IRibbonControl control)
	{
		return ExcelAddIn1.Keyboard.Shortcuts.SuperTip(control.Id);
	}

	public bool GetAutoColorState(IRibbonControl control)
	{
		return KH.A.AutoColorOnEntry;
	}

	public bool GetMaximizedState(IRibbonControl control)
	{
		return clsRibbon.GetMaximizedState();
	}

	public bool GetTranslateState(IRibbonControl control)
	{
		return Translate.IsTranslating();
	}

	public string MenuGeneral(IRibbonControl control)
	{
		return clsRibbon.MenuGeneral();
	}

	public string MenuCurrency(IRibbonControl control)
	{
		return clsRibbon.MenuCurrency();
	}

	public string MenuPercent(IRibbonControl control)
	{
		return clsRibbon.MenuPercent();
	}

	public string MenuMultiple(IRibbonControl control)
	{
		return clsRibbon.MenuMultiple();
	}

	public string MenuDate(IRibbonControl control)
	{
		return clsRibbon.MenuDate();
	}

	public string MenuBinary(IRibbonControl control)
	{
		return clsRibbon.MenuBinary();
	}

	public string MenuRatio(IRibbonControl control)
	{
		return clsRibbon.MenuRatio();
	}

	public string MenuFontStyle(IRibbonControl control)
	{
		return clsRibbon.MenuFontStyle();
	}

	public string MenuFontSize(IRibbonControl control)
	{
		return clsRibbon.MenuFontSize();
	}

	public string MenuCustom(IRibbonControl control)
	{
		return clsRibbon.MenuCustom(control);
	}

	public string MenuHeight(IRibbonControl control)
	{
		return clsRibbon.MenuHeight();
	}

	public string MenuWidth(IRibbonControl control)
	{
		return clsRibbon.MenuWidth();
	}

	public string BorderColorCycleMenu(IRibbonControl control)
	{
		return clsRibbon.BorderColorMenu();
	}

	public string ChartColorCycleMenu(IRibbonControl control)
	{
		return clsRibbon.ChartColorMenu();
	}

	public string FontColorCycleMenu(IRibbonControl control)
	{
		return clsRibbon.FontColorMenu();
	}

	public string FillColorCycleMenu(IRibbonControl control)
	{
		return clsRibbon.FillColorMenu();
	}

	public void AutomaticFont(IRibbonControl control)
	{
		FontColor.Automatic();
	}

	public void NoFill(IRibbonControl control)
	{
		FillColor.None();
	}

	public void NoBorder(IRibbonControl control)
	{
		ExcelAddIn1.Format.Borders.NoBorder();
	}

	public string AutoColorMenu(IRibbonControl control)
	{
		return clsRibbon.AutoColorMenu();
	}

	public string CellsToStandardSizeMenu(IRibbonControl control)
	{
		return clsRibbon.CellsToStandardSizeMenu();
	}

	public string ChartResizeMenu(IRibbonControl control)
	{
		return clsRibbon.ChartResizeMenu();
	}

	public string ShowGuideMenu(IRibbonControl control)
	{
		return clsRibbon.ShowGuideMenu();
	}

	public void ActivateMacabacus(ref IRibbonControl control)
	{
		Microsoft.Office.Interop.Excel.Application application = null;
		bool? flag = null;
		try
		{
			application = MH.A.Application;
			flag = application.DisplayAlerts;
			application.DisplayAlerts = false;
			Licensing.Activate();
		}
		finally
		{
			if (flag.HasValue)
			{
				application.DisplayAlerts = flag.Value;
			}
			application = null;
		}
	}

	public bool IsGroupVisible(IRibbonControl control)
	{
		return true;
	}

	public void PurchaseMacabacus(ref IRibbonControl control)
	{
		Ribbon.PurchaseMacabacus();
	}

	public bool ShowExpiredNotice(IRibbonControl control)
	{
		return Ribbon.ShowLicenseGroup();
	}

	public bool ShowNewerVersionNotice(IRibbonControl control)
	{
		return clsUpdate.ShowNewerVersionNotice(KH.A);
	}

	public void DownloadUpdate(ref IRibbonControl control)
	{
		clsUpdate.DownloadUpdate(KH.A);
	}

	public void DismissUpdate(ref IRibbonControl control)
	{
		clsUpdate.DismissUpdate(KH.A);
	}

	public string UpdateLabel(IRibbonControl control)
	{
		return clsUpdate.NewerVersionLabel();
	}

	public bool UndoEnabled(IRibbonControl control)
	{
		return ExcelAddIn1.UndoRedo.Core.UndoButtonEnabled();
	}

	public bool RedoEnabled(IRibbonControl control)
	{
		return ExcelAddIn1.UndoRedo.Core.RedoButtonEnabled();
	}

	public void Undo(ref IRibbonControl control)
	{
		ExcelAddIn1.UndoRedo.Core.Undo();
	}

	public void Redo(ref IRibbonControl control)
	{
		ExcelAddIn1.UndoRedo.Core.Redo();
	}

	public void DoGeneral(ref IRibbonControl control)
	{
		NumberFormat.DoGeneral(control);
	}

	public void DoCurrency(ref IRibbonControl control)
	{
		NumberFormat.DoCurrency(control);
	}

	public void DoPercent(ref IRibbonControl control)
	{
		NumberFormat.DoPercent(control);
	}

	public void DoMultiple(ref IRibbonControl control)
	{
		NumberFormat.DoMultiple(control);
	}

	public void DoDate(ref IRibbonControl control)
	{
		NumberFormat.DoDate(control);
	}

	public void DoBinary(ref IRibbonControl control)
	{
		NumberFormat.DoBinary(control);
	}

	public void DoRatio(ref IRibbonControl control)
	{
		NumberFormat.DoRatio(control);
	}

	public void IncrDecimal(ref IRibbonControl control)
	{
		Decimals.Increase();
	}

	public void DecrDecimal(ref IRibbonControl control)
	{
		Decimals.Decrease();
	}

	public void BlueBlackToggle(ref IRibbonControl control)
	{
		FontColor.BlueBlackToggle();
	}

	public void DoChartColor(ref IRibbonControl control)
	{
		CycleColor.DoChartColor(control);
	}

	public void ShadeRowsColumns(ref IRibbonControl control)
	{
		AlternateShading.ShadeRowsColumns(control.Tag);
	}

	public void DynamicAutoColorToggle(ref IRibbonControl control, bool pressed)
	{
		clsRibbon.DynamicAutoColorToggle();
	}

	public void AutoColorSelection(ref IRibbonControl control)
	{
		AutoColor.Selection();
	}

	public void AutoColorSheet(ref IRibbonControl control)
	{
		AutoColor.Worksheet();
	}

	public void AutoColorWorkbook(ref IRibbonControl control)
	{
		AutoColor.Workbook();
	}

	public void DoAlignHorizontal(ref IRibbonControl control)
	{
		Alignment.DoAlignHorizontal(control);
	}

	public void DoAlignVertical(ref IRibbonControl control)
	{
		Alignment.DoAlignVertical(control);
	}

	public void LeftIndent(ref IRibbonControl control)
	{
		Indent.Left();
	}

	public void RightIndent(ref IRibbonControl control)
	{
		Indent.Right();
	}

	public void IncrFont(ref IRibbonControl control)
	{
		FontSize.Increase();
	}

	public void DecrFont(ref IRibbonControl control)
	{
		FontSize.Decrease();
	}

	public void DoFontSize(ref IRibbonControl control)
	{
		FontSize.DoFontSize(control);
	}

	public void DoFontStyle(ref IRibbonControl control)
	{
		ExcelAddIn1.Format.FontStyle.DoFontStyle(control);
	}

	public void BorderTop(ref IRibbonControl control)
	{
		ExcelAddIn1.Format.Borders.CycleTop();
	}

	public void BorderBottom(ref IRibbonControl control)
	{
		ExcelAddIn1.Format.Borders.CycleBottom();
	}

	public void BorderLeft(ref IRibbonControl control)
	{
		ExcelAddIn1.Format.Borders.CycleLeft();
	}

	public void BorderRight(ref IRibbonControl control)
	{
		ExcelAddIn1.Format.Borders.CycleRight();
	}

	public void BorderOutline(ref IRibbonControl control)
	{
		ExcelAddIn1.Format.Borders.Outside();
	}

	public void BorderInside(ref IRibbonControl control)
	{
		ExcelAddIn1.Format.Borders.Inside();
	}

	public void BorderNone(ref IRibbonControl control)
	{
		ExcelAddIn1.Format.Borders.None();
	}

	public void BulletedList(ref IRibbonControl control)
	{
		Lists.Bullets();
	}

	public void DashedList(ref IRibbonControl control)
	{
		Lists.Dashes();
	}

	public void NumberedList(ref IRibbonControl control)
	{
		Lists.Numbers();
	}

	public void LetterUpperList(ref IRibbonControl control)
	{
		Lists.LettersUpper();
	}

	public void LetterLowerList(ref IRibbonControl control)
	{
		Lists.LettersLower();
	}

	public void RomanUpperList(ref IRibbonControl control)
	{
		Lists.RomanUpper();
	}

	public void RomanLowerList(ref IRibbonControl control)
	{
		Lists.RomanLower();
	}

	public void NoList(ref IRibbonControl control)
	{
		Lists.None();
	}

	public void DoFootnote(ref IRibbonControl control)
	{
		Footnotes.DoFootnote(control);
	}

	public void FootnoteToggle(ref IRibbonControl control)
	{
		Footnotes.Toggle();
	}

	public void FootnotesShow(ref IRibbonControl control)
	{
		Footnotes.Show();
	}

	public void FootnotesHide(ref IRibbonControl control)
	{
		Footnotes.Hide();
	}

	public void FootnoteSequence(ref IRibbonControl control)
	{
		Footnotes.CheckSequence();
	}

	public void PaintbrushCapture(ref IRibbonControl control)
	{
		Paintbrush.Capture();
	}

	public void PaintbrushApply(ref IRibbonControl control)
	{
		Paintbrush.Apply();
	}

	public void PaintbrushClear(ref IRibbonControl control)
	{
		Paintbrush.Clear();
	}

	public void DoUnderline(ref IRibbonControl control)
	{
		Underline.DoUnderline(control);
	}

	public void DoCycleCase(ref IRibbonControl control)
	{
		Cases.DoCycleCase(control);
	}

	public void LeaderDots(ref IRibbonControl control)
	{
		ExcelAddIn1.Format.Miscellaneous.LeaderDots();
	}

	public void SumBar(ref IRibbonControl control)
	{
		ExcelAddIn1.Format.SumBar.Toggle();
	}

	public void WrapText(ref IRibbonControl control)
	{
		ExcelAddIn1.Format.Miscellaneous.WrapText();
	}

	public void ToggleFormatPainter(ref IRibbonControl control, bool pressed)
	{
		ExcelAddIn1.FormatPainter.Pane.Toggle(pressed);
	}

	public bool IsFormatPainterOpen(IRibbonControl control)
	{
		return ExcelAddIn1.FormatPainter.Pane.IsVisible();
	}

	public void CopyProperties(IRibbonControl control)
	{
		Ribbon.Copy();
	}

	public void ApplyChartSize(IRibbonControl control)
	{
		Ribbon.ChartSize();
	}

	public void ApplyPlotSize(IRibbonControl control)
	{
		Ribbon.PlotSize();
	}

	public void ApplyPlotPosition(IRibbonControl control)
	{
		Ribbon.PlotPosition();
	}

	public void ApplyLayout(IRibbonControl control)
	{
		Ribbon.Layout();
	}

	public void ShiftDecimalLeft(ref IRibbonControl control)
	{
		ShiftDecimal.Left();
	}

	public void ShiftDecimalRight(ref IRibbonControl control)
	{
		ShiftDecimal.Right();
	}

	public void IncrTableSize(ref IRibbonControl control)
	{
		ExcelAddIn1.Format.Tables.IncreaseSize();
	}

	public void DecrTableSize(ref IRibbonControl control)
	{
		ExcelAddIn1.Format.Tables.DecreaseSize();
	}

	public void DoRowHeight(ref IRibbonControl control)
	{
		CellSize.DoRowHeight(control);
	}

	public void DoColumnWidth(ref IRibbonControl control)
	{
		CellSize.DoColumnWidth(control);
	}

	public void AutoFitRow(ref IRibbonControl control)
	{
		AutoFit.Height();
	}

	public void AutoFitCol(ref IRibbonControl control)
	{
		AutoFit.Width();
	}

	public void ConformSize(ref IRibbonControl control)
	{
		CellSize.ConformSize(Conversions.ToInteger(control.Tag));
	}

	public void ConformPowerPoint(ref IRibbonControl control)
	{
		CellSize.ConformPowerPoint();
	}

	public void ConformWord(ref IRibbonControl control)
	{
		CellSize.ConformWord();
	}

	public void ShowGuide(ref IRibbonControl control)
	{
		Guides.Show(Conversions.ToInteger(control.Tag));
	}

	public void ShowGuideXL(ref IRibbonControl control)
	{
		Guides.ShowExcel();
	}

	public void ShowGuidePP(ref IRibbonControl control)
	{
		Guides.ShowPowerPoint();
	}

	public void ShowGuideWD(ref IRibbonControl control)
	{
		Guides.ShowWord();
	}

	public void RemoveGuides(ref IRibbonControl control)
	{
		Guides.Remove();
	}

	public void InsertRow(ref IRibbonControl control)
	{
		Insert.Row();
	}

	public void InsertCol(ref IRibbonControl control)
	{
		Insert.Column();
	}

	public void DeleteRow(ref IRibbonControl control)
	{
		Delete.Row();
	}

	public void DeleteCol(ref IRibbonControl control)
	{
		Delete.Column();
	}

	public void DeleteBlankRows(ref IRibbonControl control)
	{
		Delete.BlankRows();
	}

	public void DeleteBlankColumns(ref IRibbonControl control)
	{
		Delete.BlankColumns();
	}

	public void GroupRow(ref IRibbonControl control)
	{
		Group.Rows();
	}

	public void GroupCol(ref IRibbonControl control)
	{
		Group.Columns();
	}

	public void UngroupRow(ref IRibbonControl control)
	{
		Ungroup.Rows();
	}

	public void UngroupCol(ref IRibbonControl control)
	{
		Ungroup.Columns();
	}

	public void HideRow(ref IRibbonControl control)
	{
		Hide.Rows();
	}

	public void HideCol(ref IRibbonControl control)
	{
		Hide.Columns();
	}

	public void UnhideRow(ref IRibbonControl control)
	{
		Unhide.Rows();
	}

	public void UnhideCol(ref IRibbonControl control)
	{
		Unhide.Columns();
	}

	public void ProperHide(ref IRibbonControl control)
	{
		Hide.ProperHide();
	}

	public void ExpandRows(ref IRibbonControl control)
	{
		Expand.Rows();
	}

	public void ExpandCols(ref IRibbonControl control)
	{
		Expand.Columns();
	}

	public void CollapseRows(ref IRibbonControl control)
	{
		Collapse.Rows();
	}

	public void CollapseCols(ref IRibbonControl control)
	{
		Collapse.Columns();
	}

	public void RowColCopy(ref IRibbonControl control)
	{
		Info.Copy();
	}

	public void RowColPaste(ref IRibbonControl control)
	{
		Info.Paste();
	}

	public void ModifyRows(ref IRibbonControl control)
	{
		BatchModify.Rows();
	}

	public void ModifyCols(ref IRibbonControl control)
	{
		BatchModify.Columns();
	}

	public void ReverseColumns(ref IRibbonControl control)
	{
		Reverse.Columns();
	}

	public void ReverseRows(ref IRibbonControl control)
	{
		Reverse.Rows();
	}

	public void TraceIn(ref IRibbonControl control)
	{
		ExcelAddIn1.Audit.TraceDialogs.Precedents.Dialog.Show();
	}

	public void TraceOut(ref IRibbonControl control)
	{
		ExcelAddIn1.Audit.TraceDialogs.Dependents.Dialog.Show();
	}

	public void LastAuditedCell(ref IRibbonControl control)
	{
		ExcelAddIn1.Audit.TraceDialogs.Base.GoToLastAuditedCell();
	}

	public void AutoPrecedentsToggle(ref IRibbonControl control, bool pressed)
	{
		AutoTrace.TogglePrecedents();
	}

	public void AutoDependentsToggle(ref IRibbonControl control, bool pressed)
	{
		AutoTrace.ToggleDependents();
	}

	public void ShowAllPrecedents(ref IRibbonControl control)
	{
		TraceAll.Precedents();
	}

	public void ShowAllDependents(ref IRibbonControl control)
	{
		TraceAll.Dependents();
	}

	public void ClearArrows(ref IRibbonControl control)
	{
		Arrows.Clear();
	}

	public bool GetTracePre(IRibbonControl control)
	{
		return K.Settings.AutoTracePrecedents;
	}

	public bool GetTraceDep(IRibbonControl control)
	{
		return K.Settings.AutoTraceDependents;
	}

	public void DependencyHeatmap(ref IRibbonControl control)
	{
		DependencyDensity.Apply();
	}

	public void MagnitudeHeatmap(ref IRibbonControl control)
	{
		MagnitudeMap.Apply();
	}

	public void FunctionalMap(ref IRibbonControl control)
	{
		ExcelAddIn1.Audit.Visualizations.FunctionalMap.Apply();
	}

	public void FormulaFlow(ref IRibbonControl control)
	{
		ExcelAddIn1.Audit.Visualizations.FormulaFlow.Apply();
	}

	public void ClearVisualizations(ref IRibbonControl control)
	{
		Common.ClearVisualizations();
	}

	public void FormulaReport(ref IRibbonControl control)
	{
		ExcelAddIn1.Audit.Visualizations.FormulaReport.Create();
	}

	public void UniformRange(ref IRibbonControl control)
	{
		Uniformulas.Apply();
	}

	public void NextWatch(ref IRibbonControl control)
	{
		ExcelAddIn1.Audit.Watches.GoToNext();
	}

	public void PrevWatch(ref IRibbonControl control)
	{
		ExcelAddIn1.Audit.Watches.GoToPrevious();
	}

	public void AddWatch(ref IRibbonControl control)
	{
		ExcelAddIn1.Audit.Watches.Add();
	}

	public void RemoveWatch(ref IRibbonControl control)
	{
		ExcelAddIn1.Audit.Watches.Remove();
	}

	public void ClearWatches(ref IRibbonControl control)
	{
		ExcelAddIn1.Audit.Watches.Clear();
	}

	public void WatchWindow(ref IRibbonControl control)
	{
		ExcelAddIn1.Audit.Watches.WatchWindow();
	}

	public void DiscussPane(ref IRibbonControl control, bool pressed)
	{
		clsDiscuss.DiscussPaneToggle(pressed);
	}

	public bool IsDiscussPaneOpen(IRibbonControl control)
	{
		return clsDiscuss.IsDiscussPaneOpen();
	}

	public void ProCopyRight(ref IRibbonControl control)
	{
		FastFill.Right();
	}

	public void ProCopyDown(ref IRibbonControl control)
	{
		FastFill.Down();
	}

	public void CheckForErrors(ref IRibbonControl control)
	{
		ErrorWrap.Toggle();
	}

	public void CleanFormula(ref IRibbonControl control)
	{
		Clean.Selection();
	}

	public void PrependEquals(ref IRibbonControl control)
	{
		ExcelAddIn1.Formulas.PrependEquals.Go();
	}

	public void CommentFormula(ref IRibbonControl control)
	{
		ExcelAddIn1.Formulas.Comment.Cells();
	}

	public void WrapParentheses(ref IRibbonControl control)
	{
		ExcelAddIn1.Formulas.WrapParentheses.Go();
	}

	public void ConvFormula(ref IRibbonControl control)
	{
		Anchor.Cycle();
	}

	public void UnapplyNames(ref IRibbonControl control)
	{
		ExcelAddIn1.Formulas.Names.Unapply();
	}

	public void AnchorFormula(ref IRibbonControl control)
	{
		Anchor.Convert(Conversions.ToInteger(control.Tag));
	}

	public void Flatten(ref IRibbonControl control)
	{
		ExcelAddIn1.Formulas.Flatten.Selection();
	}

	public void IsolateSheets(ref IRibbonControl control)
	{
		ExcelAddIn1.Formulas.Flatten.IsolateSheets();
	}

	public void FlattenFunction(ref IRibbonControl control)
	{
		ExcelAddIn1.Formulas.Flatten.FlattenFunction();
	}

	public void FlipSign(ref IRibbonControl control)
	{
		ExcelAddIn1.Formulas.FlipSign.Go();
	}

	public void InsertSymbol(ref IRibbonControl control)
	{
		Symbols.Insert(Conversions.ToInteger(control.Tag));
	}

	public void Untranspose(ref IRibbonControl control)
	{
		ExcelAddIn1.Formulas.Untranspose.Go();
	}

	public void UnIndirect(ref IRibbonControl control)
	{
		Evaluate.Indirect();
	}

	public void UnChoose(ref IRibbonControl control)
	{
		Evaluate.Choose();
	}

	public void UnOffset(ref IRibbonControl control)
	{
		Evaluate.Offset();
	}

	public void UnVLookup(ref IRibbonControl control)
	{
		Evaluate.VLookup();
	}

	public void UnHLookup(ref IRibbonControl control)
	{
		Evaluate.HLookup();
	}

	public void UnXLookup(ref IRibbonControl control)
	{
		Evaluate.XLookup();
	}

	public void UnIndexMatch(ref IRibbonControl control)
	{
		Evaluate.IndexMatch();
	}

	public void UnIf(ref IRibbonControl control)
	{
		Evaluate.IfFunction();
	}

	public void UnMin(ref IRibbonControl control)
	{
		Evaluate.Min();
	}

	public void UnSumIf(ref IRibbonControl control)
	{
		Evaluate.SumIf();
	}

	public void UnSumIfs(ref IRibbonControl control)
	{
		Evaluate.SumIfs();
	}

	public void UnMax(ref IRibbonControl control)
	{
		Evaluate.Max();
	}

	public void SimplifyFormula(ref IRibbonControl control)
	{
		Evaluate.SimplifyFormula();
	}

	public void SimplifyIndirect(ref IRibbonControl control, bool pressed)
	{
		K.Settings.SimplifyIndirect = pressed;
	}

	public void SimplifyChoose(ref IRibbonControl control, bool pressed)
	{
		K.Settings.SimplifyChoose = pressed;
	}

	public void SimplifyOffset(ref IRibbonControl control, bool pressed)
	{
		K.Settings.SimplifyOffset = pressed;
	}

	public void SimplifyHlookup(ref IRibbonControl control, bool pressed)
	{
		K.Settings.SimplifyHlookup = pressed;
	}

	public void SimplifyVlookup(ref IRibbonControl control, bool pressed)
	{
		K.Settings.SimplifyVlookup = pressed;
	}

	public void SimplifyXlookup(ref IRibbonControl control, bool pressed)
	{
		K.Settings.SimplifyXlookup = pressed;
	}

	public void SimplifyIndexMatch(ref IRibbonControl control, bool pressed)
	{
		K.Settings.SimplifyIndexMatch = pressed;
	}

	public void SimplifyIf(ref IRibbonControl control, bool pressed)
	{
		K.Settings.SimplifyIf = pressed;
	}

	public void SimplifyMin(ref IRibbonControl control, bool pressed)
	{
		K.Settings.SimplifyMin = pressed;
	}

	public void SimplifyMax(ref IRibbonControl control, bool pressed)
	{
		K.Settings.SimplifyMax = pressed;
	}

	public void SimplifySumIf(ref IRibbonControl control, bool pressed)
	{
		K.Settings.SimplifySumIf = pressed;
	}

	public void SimplifySumIfs(ref IRibbonControl control, bool pressed)
	{
		K.Settings.SimplifySumIfs = pressed;
	}

	public bool SimplifyIndirectEnabled(IRibbonControl control)
	{
		return K.Settings.SimplifyIndirect;
	}

	public bool SimplifyChooseEnabled(IRibbonControl control)
	{
		return K.Settings.SimplifyChoose;
	}

	public bool SimplifyOffsetEnabled(IRibbonControl control)
	{
		return K.Settings.SimplifyOffset;
	}

	public bool SimplifyHlookupEnabled(IRibbonControl control)
	{
		return K.Settings.SimplifyHlookup;
	}

	public bool SimplifyVlookupEnabled(IRibbonControl control)
	{
		return K.Settings.SimplifyVlookup;
	}

	public bool SimplifyXlookupEnabled(IRibbonControl control)
	{
		return K.Settings.SimplifyXlookup;
	}

	public bool SimplifyIndexMatchEnabled(IRibbonControl control)
	{
		return K.Settings.SimplifyIndexMatch;
	}

	public bool SimplifyIfEnabled(IRibbonControl control)
	{
		return K.Settings.SimplifyIf;
	}

	public bool SimplifyMinEnabled(IRibbonControl control)
	{
		return K.Settings.SimplifyMin;
	}

	public bool SimplifyMaxEnabled(IRibbonControl control)
	{
		return K.Settings.SimplifyMax;
	}

	public bool SimplifySumIfEnabled(IRibbonControl control)
	{
		return K.Settings.SimplifySumIf;
	}

	public bool SimplifySumIfsEnabled(IRibbonControl control)
	{
		return K.Settings.SimplifySumIfs;
	}

	public void ValidationToggle(ref IRibbonControl control)
	{
		ExcelAddIn1.Model.Validation.Toggle();
	}

	public void ValidationNumber(ref IRibbonControl control)
	{
		ExcelAddIn1.Model.Validation.Number();
	}

	public void ValidationDate(ref IRibbonControl control)
	{
		ExcelAddIn1.Model.Validation.DateInput();
	}

	public void ValidationText(ref IRibbonControl control)
	{
		ExcelAddIn1.Model.Validation.Text();
	}

	public void ValidationZeroOrLarger(ref IRibbonControl control)
	{
		ExcelAddIn1.Model.Validation.ZeroOrLarger();
	}

	public void ValidationZeroOrSmaller(ref IRibbonControl control)
	{
		ExcelAddIn1.Model.Validation.ZeroOrSmaller();
	}

	public void ValidationPositivePercent(ref IRibbonControl control)
	{
		ExcelAddIn1.Model.Validation.PositivePercent();
	}

	public void ValidationAnyPercent(ref IRibbonControl control)
	{
		ExcelAddIn1.Model.Validation.AnyPercent();
	}

	public void ValidationClear(ref IRibbonControl control)
	{
		ExcelAddIn1.Model.Validation.Clear();
	}

	public void PasteNumFormat(ref IRibbonControl control)
	{
		Paste.NumberFormats();
	}

	public void PasteExact(ref IRibbonControl control)
	{
		Paste.Exact(trans: false);
	}

	public void PasteDuplicate(ref IRibbonControl control)
	{
		Paste.Duplicate();
	}

	public void PasteTranspose(ref IRibbonControl control)
	{
		Paste.Transpose();
	}

	public void PasteLinks(ref IRibbonControl control)
	{
		Paste.Links();
	}

	public void PasteInsert(ref IRibbonControl control)
	{
		Paste.Insert();
	}

	public void SummarizeFormula(ref IRibbonControl control)
	{
		Summarize.Go();
	}

	public void QuickCAGR(ref IRibbonControl control)
	{
		QuickCagr.Add();
	}

	public void GrowthDriver(ref IRibbonControl control)
	{
		ExcelAddIn1.Model.GrowthDriver.Add();
	}

	public void ReplicateModule(ref IRibbonControl control)
	{
		ExcelAddIn1.Model.ReplicateModule.Go();
	}

	public void ContributionAnalysis(ref IRibbonControl control)
	{
		ExcelAddIn1.Model.ContributionAnalysis.Go();
	}

	public void InsertSummaryStats(ref IRibbonControl control)
	{
		SummaryStats.Add();
	}

	public void AddScenarios(ref IRibbonControl control)
	{
		Scenarios.Add();
	}

	public void AutoFillDates(ref IRibbonControl control)
	{
		AutoFill.Dates();
	}

	public void DoFillDates(ref IRibbonControl control)
	{
		AutoFill.Dates(control);
	}

	public void ToggleTranslate(ref IRibbonControl control, bool pressed)
	{
		Translate.ToggleTranslate();
	}

	public void PrepareToShare(ref IRibbonControl control, bool pressed)
	{
		ExcelAddIn1.Publishing.Share.Pane.Toggle(pressed);
	}

	public bool IsPrepareToShareOpen(IRibbonControl control)
	{
		return ExcelAddIn1.Publishing.Share.Pane.IsVisible();
	}

	public void SanitizeReplaceFormulas(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.FlattenWorkbookPublic();
	}

	public void SanitizeColorFontsBlack(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.ColorFontsBlackPublic();
	}

	public void SanitizeRemoveCellComments(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.DeleteCommentsPublic();
	}

	public void SanitizeRemoveNames(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.RemoveNames();
	}

	public void SanitizeRemoveHiddenSheets(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.DeleteHiddenSheetsPublic();
	}

	public void SanitizeDeleteHiddenRowsCols(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.DeleteHiddenRowsColsPublic();
	}

	public void SanitizeCollapseGroupedRowsCols(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.CollapseGroupedRowsColsPublic();
	}

	public void SanitizeRemoveCharts(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.DeleteAllCharts();
	}

	public void SanitizeRemoveWatches(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.DeleteWatchesPublic();
	}

	public void SanitizeRemoveInk(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.DeleteInkPublic();
	}

	public void SanitizeCheckFormulaErrors(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.CheckFormulaErrors();
	}

	public void SanitizeResetPrintAreas(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.ResetPrintAreasPublic();
	}

	public void SanitizeHideGridlines(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.HideGridlinesPublic();
	}

	public void SanitizeZoomTo100(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.ZoomTo100Public();
	}

	public void SanitizeReturnToCellA1(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.ReturnToCellA1Public();
	}

	public void SanitizeCleanCells(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.CleanCellsPublic();
	}

	public void SanitizeBreakHyperlinks(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.BreakHyperlinksPublic();
	}

	public void VeryHideSheets(ref IRibbonControl control)
	{
		ExcelAddIn1.Publishing.Share.Base.VeryHidePublic();
	}

	public void EmailDocument(IRibbonControl control)
	{
		Send.ShowDialog();
	}

	public void PdfToFolder(ref IRibbonControl control)
	{
		Pdf.ToFolder();
	}

	public void FirstSheet(ref IRibbonControl control)
	{
		ExcelAddIn1.Sheets.Navigate.A();
	}

	public void LastSheet(ref IRibbonControl control)
	{
		ExcelAddIn1.Sheets.Navigate.B();
	}

	public void NextSheet(ref IRibbonControl control)
	{
		ExcelAddIn1.Sheets.Navigate.C();
	}

	public void PrevSheet(ref IRibbonControl control)
	{
		ExcelAddIn1.Sheets.Navigate.D();
	}

	public void SheetActivate(ref IRibbonControl control)
	{
		ExcelAddIn1.Sheets.Navigate.E();
	}

	public void SheetMoveLeft(IRibbonControl control)
	{
		Move.A();
	}

	public void SheetMoveRight(IRibbonControl control)
	{
		Move.B();
	}

	public void GoToMin(ref IRibbonControl control)
	{
		MinMax.A();
	}

	public void GoToMax(ref IRibbonControl control)
	{
		MinMax.B();
	}

	public void BurySelectedSheets(ref IRibbonControl control)
	{
		ExcelAddIn1.Sheets.Visibility.A();
	}

	public void BuryHiddenSheets(ref IRibbonControl control)
	{
		ExcelAddIn1.Sheets.Visibility.B();
	}

	public void DigUpBuriedSheets(ref IRibbonControl control)
	{
		ExcelAddIn1.Sheets.Visibility.C();
	}

	public void BackstageShow(object contextObject)
	{
		KH.A = true;
	}

	public void BackstageHide(object contextObject)
	{
		KH.A = false;
	}

	public void AuditCheck(ref IRibbonControl control, bool pressed)
	{
		if (AuditCheckEnabled(control))
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					ExcelAddIn1.Audit.Check.UI.Pane.Toggle(pressed);
					return;
				}
			}
		}
		ExcelAddIn1.Audit.Check.UI.Pane.B();
	}

	public bool AuditCheckShown(IRibbonControl control)
	{
		return ExcelAddIn1.Audit.Check.UI.Pane.IsVisible();
	}

	public bool AuditCheckEnabled(IRibbonControl control)
	{
		if (IsNotProtectedView(control))
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					int result;
					if (wpfAudit.InstanceRunningAnalysis != null)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							break;
						}
						result = ((!ExcelAddIn1.Audit.Check.UI.Pane.IsVisible()) ? 1 : 0);
					}
					else
					{
						result = 1;
					}
					return (byte)result != 0;
				}
				}
			}
		}
		return false;
	}

	public void SuperFind(ref IRibbonControl control, bool pressed)
	{
		ExcelAddIn1.SuperFind2.UI.Pane.Toggle(pressed);
	}

	public bool SuperFindShown(IRibbonControl control)
	{
		return ExcelAddIn1.SuperFind2.UI.Pane.IsVisible();
	}

	public void ZoomIn(ref IRibbonControl control)
	{
		Zoom.ZoomIn();
	}

	public void ZoomOut(ref IRibbonControl control)
	{
		Zoom.ZoomOut();
	}

	public void ToggleGridlines(ref IRibbonControl control)
	{
		ExcelAddIn1.View.Gridlines.Toggle();
	}

	public void HidePageBreaks(ref IRibbonControl control)
	{
		PrintAreas.HidePageBreaks();
	}

	public void SmartPrintArea(ref IRibbonControl control)
	{
		PrintAreas.SmartPrintArea();
	}

	public void SetPrintAreas(ref IRibbonControl control)
	{
		PrintAreas.SetPrintAreas();
	}

	public void RemovePrintAreas(ref IRibbonControl control)
	{
		PrintAreas.RemovePrintAreas();
	}

	public void FreezePanes(ref IRibbonControl control)
	{
		ExcelAddIn1.View.FreezePanes.Freeze();
	}

	public void UnfreezePanes(ref IRibbonControl control)
	{
		ExcelAddIn1.View.FreezePanes.Unfreeze();
	}

	public void Split(ref IRibbonControl control)
	{
		Splits.Split();
	}

	public void Unsplit(ref IRibbonControl control)
	{
		Splits.Unsplit();
	}

	public void LockScrollArea(ref IRibbonControl control)
	{
		ScrollArea.Lock();
	}

	public void UnlockScrollArea(ref IRibbonControl control)
	{
		ScrollArea.Unlock();
	}

	public void NavAidToggle(ref IRibbonControl control, bool pressed)
	{
		NavAid.Toggle(pressed);
	}

	public bool NavAidEnabled(IRibbonControl control)
	{
		return NavAid.Enabled;
	}

	public void MaximizeWorkspace(ref IRibbonControl control, bool pressed)
	{
		Workspace.Maximize(pressed);
	}

	public void FixNotes(ref IRibbonControl control)
	{
		Fix.NotesInSelection();
	}

	public void DeleteEmptyNotes(ref IRibbonControl control)
	{
		CleanUp.DeleteEmptyNotes();
	}

	public void CommentDelete(ref IRibbonControl control)
	{
		CleanUp.Delete();
	}

	public void DeleteResolvedComments(ref IRibbonControl control)
	{
		ThreadedComments.DeleteResolved();
	}

	public void ResolveComments(ref IRibbonControl control)
	{
		ThreadedComments.Resolve();
	}

	public void ReopenComments(ref IRibbonControl control)
	{
		ThreadedComments.Reopen();
	}

	public void ConvertNoteToComment(ref IRibbonControl control)
	{
		ThreadedComments.ConvertFromNote();
	}

	public void RemoveAuthor(ref IRibbonControl control)
	{
		Author.Remove();
	}

	public void ChangeCommentAuthor(ref IRibbonControl control)
	{
		Author.Change();
	}

	public void QuickSave(ref IRibbonControl control)
	{
		ExcelAddIn1.Workbook.QuickSave.Save();
	}

	public void QuickSaveAll(ref IRibbonControl control)
	{
		ExcelAddIn1.Workbook.QuickSave.SaveAll();
	}

	public void QuickSaveAs(ref IRibbonControl control)
	{
		ExcelAddIn1.Workbook.QuickSave.SaveAs();
	}

	public void QuickSaveUp(ref IRibbonControl control)
	{
		ExcelAddIn1.Workbook.QuickSave.SaveUp();
	}

	public void QuickSaveDown(ref IRibbonControl control)
	{
		ExcelAddIn1.Workbook.QuickSave.SaveDown();
	}

	public void MergeFiles(ref IRibbonControl control)
	{
		ExcelAddIn1.Workbook.Merge.Dialog.Show();
	}

	public void Reopen(ref IRibbonControl control)
	{
		ExcelAddIn1.Workbook.Miscellaneous.Reopen();
	}

	public void ShowInFolder(ref IRibbonControl control)
	{
		ExcelAddIn1.Workbook.Miscellaneous.OpenFolder();
	}

	public void CopyPath(ref IRibbonControl control)
	{
		ExcelAddIn1.Workbook.Miscellaneous.CopyPath();
	}

	public void Duplicate(ref IRibbonControl control)
	{
		ExcelAddIn1.Workbook.Miscellaneous.Duplicate();
	}

	public void CloseOtherWorkbooks(ref IRibbonControl control)
	{
		ExcelAddIn1.Workbook.Miscellaneous.CloseOthers();
	}

	public void NameManager(ref IRibbonControl control)
	{
		Optimize.NameScrubber();
	}

	public void CleanUpUsedRanges(ref IRibbonControl control)
	{
		Optimize.CleanUpUsedRanges();
	}

	public void BatchRemoveStyles(ref IRibbonControl control)
	{
		Optimize.StyleScrubber();
	}

	public void DeleteUnusedCustomNumberFormats(ref IRibbonControl control)
	{
		Optimize.DeleteUnusedCustomNumberFormats();
	}

	public void ClearConstants(ref IRibbonControl control)
	{
		ExcelAddIn1.Workbook.Miscellaneous.ClearConstants();
	}

	public void Shortcuts(ref IRibbonControl control)
	{
	}

	public void ChartPlotSize(ref IRibbonControl control)
	{
		ChartAndPlotSize.ShowDialog();
	}

	public void ChartSize(ref IRibbonControl control)
	{
		ResizeTo.StandardSize(Conversions.ToInteger(control.Tag));
	}

	public void ResizeChartToPP(ref IRibbonControl control)
	{
		ResizeTo.PowerPointSelection();
	}

	public void ResizeChartToWD(ref IRibbonControl control)
	{
		ResizeTo.WordSelection();
	}

	public void ResizeChartToXL(ref IRibbonControl control)
	{
		ResizeTo.ExcelSelection();
	}

	public void ChartGallery(ref IRibbonControl control)
	{
		ExcelAddIn1.Library2.Charts.ShowInLibrary();
	}

	public void FastFormatCharts(ref IRibbonControl control)
	{
		ExcelAddIn1.FastFormats.Charts.Apply.Selection();
	}

	public void StackedColumnChart(ref IRibbonControl control)
	{
		StackedColumn.Create();
	}

	public void WaterfallChart(ref IRibbonControl control)
	{
		Waterfall.Create();
	}

	public void StackedWaterfallChart(ref IRibbonControl control)
	{
		StackedWaterfall.Create();
	}

	public void FootballFieldChart(ref IRibbonControl control)
	{
		FootballField.Create();
	}

	public void ButterflyChart(ref IRibbonControl control)
	{
		Butterfly.Create();
	}

	public void MarimekkoChart(ref IRibbonControl control)
	{
		Marimekko.Create();
	}

	public void ScatterChart(ref IRibbonControl control)
	{
		Scatter.Create();
	}

	public void GanttChart(ref IRibbonControl control)
	{
		PriceAnnotate.Create();
	}

	public void MemorizeChart(ref IRibbonControl control)
	{
		MemorizeApply.Memorize();
	}

	public void SetToMemorized(ref IRibbonControl control)
	{
		MemorizeApply.SetToMemorized((MemorizeApply.MemorizedProperty)Conversions.ToInteger(control.Tag));
	}

	public void RecolorSeriesToDefaults(ref IRibbonControl control)
	{
		ExcelAddIn1.Charts.Recolor.SeriesToDefaults();
	}

	public void RecolorPointsToSource(ref IRibbonControl control)
	{
		ExcelAddIn1.Charts.Recolor.PointsToSource();
	}

	public void RecolorLabelsToPoints(ref IRibbonControl control)
	{
		ExcelAddIn1.Charts.Recolor.LabelsToPoints();
	}

	public void LinkFormatsToCells(ref IRibbonControl control)
	{
		ExcelAddIn1.Charts.DataLabels.LinkFormatsToCells();
	}

	public void RotateLabelsHoriz(ref IRibbonControl control)
	{
		ExcelAddIn1.Charts.DataLabels.RotateHorizontal();
	}

	public void RotateLabels90(ref IRibbonControl control)
	{
		ExcelAddIn1.Charts.DataLabels.Rotate90();
	}

	public void RotateLabels270(ref IRibbonControl control)
	{
		ExcelAddIn1.Charts.DataLabels.Rotate270();
	}

	public void RotateLabelsStacked(ref IRibbonControl control)
	{
		ExcelAddIn1.Charts.DataLabels.RotateStacked();
	}

	public void MoveDataLabels(ref IRibbonControl control, bool pressed)
	{
		ExcelAddIn1.Charts.MoveDataLabels.Pane.A(pressed);
	}

	public bool IsMoveDataLabelsOpen(IRibbonControl control)
	{
		return ExcelAddIn1.Charts.MoveDataLabels.Pane.A();
	}

	public void GrowthArrow(ref IRibbonControl control, bool pressed)
	{
		ExcelAddIn1.Charts.GrowthArrow.Pane.Toggle(pressed);
	}

	public bool IsGrowthArrowOpen(IRibbonControl control)
	{
		return ExcelAddIn1.Charts.GrowthArrow.Pane.IsVisible();
	}

	public void StackCharts(ref IRibbonControl control)
	{
		Stack.Initiate();
	}

	public void SaveChartAsPicture(ref IRibbonControl control)
	{
		SaveAsImage.Initiate();
	}

	public void ReplaceMissingLabels(ref IRibbonControl control)
	{
		ExcelAddIn1.Charts.DataLabels.ReplaceMissingLabels(control.Tag);
	}

	public void AttachLabelsToPoints(ref IRibbonControl control)
	{
		ExcelAddIn1.Charts.DataLabels.AttachLabelsToPoints();
	}

	public void StackedColumnTotals(ref IRibbonControl control)
	{
		ExcelAddIn1.Charts.StackedColumnTotals.Add();
	}

	public void StackedBarTotals(ref IRibbonControl control)
	{
		ExcelAddIn1.Charts.StackedBarTotals.Add();
	}

	public void StatLineAverage(ref IRibbonControl control)
	{
		StatLines.Average();
	}

	public void StatLineMedian(ref IRibbonControl control)
	{
		StatLines.Median();
	}

	public void StatLineValue(ref IRibbonControl control)
	{
		StatLines.Value();
	}

	public void TargetBand(ref IRibbonControl control)
	{
		ExcelAddIn1.Charts.TargetBand.Add();
	}

	public void LabelPoints(ref IRibbonControl control)
	{
		ExcelAddIn1.Charts.DataLabels.LabelPoints();
	}

	public void ChartAxisMaxFormat(ref IRibbonControl control)
	{
		YAxisFormat.ChartAxisMaxFormat(control);
	}

	public void ChartAxisFormat(ref IRibbonControl control)
	{
		YAxisFormat.ChartAxisFormat(control);
	}

	public void RescaleAxis(ref IRibbonControl control)
	{
		YAxisFormat.RescaleAxis();
	}

	public void SmartPlotOrder(ref IRibbonControl control)
	{
		PlotOrder.SmartSort();
	}

	public void ExportToPowerPoint(IRibbonControl control)
	{
		Export.A();
	}

	public void ExportToWord(IRibbonControl control)
	{
		Export.B();
	}

	public void ExportToPowerPointAsGraphic(IRibbonControl control)
	{
		Export.C();
	}

	public void ExportToPowerPointAsImage(IRibbonControl control)
	{
		Export.D();
	}

	public void ExportToPowerPointAsTable(IRibbonControl control)
	{
		Export.E();
	}

	public void ExportToPowerPointAsEmbedded(IRibbonControl control)
	{
		Export.F();
	}

	public void ExportToPowerPointAsChart(IRibbonControl control)
	{
		Export.H();
	}

	public void ExportToPowerPointAsText(IRibbonControl control)
	{
		Export.G();
	}

	public void ExportToWordAsGraphic(IRibbonControl control)
	{
		Export.I();
	}

	public void ExportToWordAsImage(IRibbonControl control)
	{
		Export.J();
	}

	public void ExportToWordAsTable(IRibbonControl control)
	{
		Export.K();
	}

	public void ExportToWordAsEmbedded(IRibbonControl control)
	{
		Export.L();
	}

	public void ExportToWordAsChart(IRibbonControl control)
	{
		Export.N();
	}

	public void ExportToWordAsText(IRibbonControl control)
	{
		Export.M();
	}

	public void CheckMatchWidth(ref IRibbonControl control, bool pressed)
	{
		K.Settings.ExportMatchDestinationWidth = pressed;
	}

	public void CheckMatchHeight(ref IRibbonControl control, bool pressed)
	{
		K.Settings.ExportMatchDestinationHeight = pressed;
	}

	public bool GetMatchWidthChecked(IRibbonControl control)
	{
		return K.Settings.ExportMatchDestinationWidth;
	}

	public bool GetMatchHeightChecked(IRibbonControl control)
	{
		return K.Settings.ExportMatchDestinationHeight;
	}

	public void PrintAreasSheetPowerPoint(IRibbonControl control)
	{
		PagesToPowerPoint.PrintAreasSelectedSheets(MH.A.Application, (Microsoft.Office.Interop.PowerPoint.Application)null);
	}

	public void PrintAreasAllPowerPoint(IRibbonControl control)
	{
		PagesToPowerPoint.PrintAreasAllSheets(MH.A.Application, (Microsoft.Office.Interop.PowerPoint.Application)null);
	}

	public void PrintAreasSheetWord(IRibbonControl control)
	{
		PagesToWord.PrintAreasSelectedSheets(MH.A.Application, (Microsoft.Office.Interop.Word.Application)null);
	}

	public void PrintAreasAllWord(IRibbonControl control)
	{
		PagesToWord.PrintAreasAllSheets(MH.A.Application, (Microsoft.Office.Interop.Word.Application)null);
	}

	public void InsertContent(ref IRibbonControl control, bool pressed)
	{
		ExcelAddIn1.Library2.UI.Pane.Toggle(pressed);
	}

	public bool ContentPaneShown(IRibbonControl control)
	{
		return ExcelAddIn1.Library2.UI.Pane.IsVisible();
	}

	public string ModelLibraryMenu(IRibbonControl control)
	{
		return Models.BuildFilesMenu();
	}

	public void OpenFile(IRibbonControl control)
	{
		Models.OpenFile(control.Tag);
	}

	public void LibraryShowAll(IRibbonControl control)
	{
		ExcelAddIn1.Library2.UI.Pane.Show(blnShapes: true, blnImages: true, blnCharts: true, blnText: true, blnTables: true, blnModels: true, blnPDFs: true);
	}

	public void LibraryShowShapes(IRibbonControl control)
	{
		ExcelAddIn1.Library2.UI.Pane.Show(blnShapes: true, blnImages: false, blnCharts: false, blnText: false, blnTables: false);
	}

	public void LibraryShowImages(IRibbonControl control)
	{
		ExcelAddIn1.Library2.UI.Pane.Show(blnShapes: false, blnImages: true, blnCharts: false, blnText: false, blnTables: false);
	}

	public void LibraryShowCharts(IRibbonControl control)
	{
		ExcelAddIn1.Library2.UI.Pane.Show(blnShapes: false, blnImages: false, blnCharts: true, blnText: false, blnTables: false);
	}

	public void LibraryShowText(IRibbonControl control)
	{
		ExcelAddIn1.Library2.UI.Pane.Show(blnShapes: false, blnImages: false, blnCharts: false, blnText: true, blnTables: false);
	}

	public void LibraryShowTables(IRibbonControl control)
	{
		ExcelAddIn1.Library2.UI.Pane.Show(blnShapes: false, blnImages: false, blnCharts: false, blnText: false, blnTables: true);
	}

	public void ManageLibraryContents(IRibbonControl control)
	{
		Admin.LibraryManager((Microsoft.Office.Interop.PowerPoint.Application)null, MH.A.Application, (Microsoft.Office.Interop.Word.Application)null);
	}

	public void LibraryVersionControl(IRibbonControl control, bool pressed)
	{
		ExcelAddIn1.Library2.Versioning.Pane.Toggle(pressed);
	}

	public bool IsLibContentPaneOpen(IRibbonControl control)
	{
		return ExcelAddIn1.Library2.Versioning.Pane.IsVisible();
	}

	public void SoftDisableToggle(IRibbonControl control)
	{
		SoftDisable.SoftToggle();
	}

	public Bitmap GetSoftDisableImage(IRibbonControl control)
	{
		return SoftDisable.GetButtonImage();
	}

	public string GetSoftDisableLabel(IRibbonControl control)
	{
		return SoftDisable.GetButtonLabel();
	}

	public void HotkeyManager(ref IRibbonControl control, bool pressed)
	{
		ShortcutManager.Toggle(pressed);
	}

	public bool IsHotkeyManagerOpen(IRibbonControl control)
	{
		return ShortcutManager.IsOpen();
	}

	public void OverrideHotkeys(IRibbonControl control)
	{
		ExcelAddIn1.Keyboard.Shortcuts.OverrideHotkeys();
	}

	public void OpenHotkeysPdf(IRibbonControl control)
	{
		ExcelAddIn1.Keyboard.Shortcuts.OpenHotkeysPdf();
	}

	public void OpenHotkeysSheet(IRibbonControl control)
	{
		ExcelAddIn1.Keyboard.Shortcuts.OpenHotkeysSheet();
	}

	public void ToggleDisabledKey(ref IRibbonControl control, bool pressed)
	{
		DisabledKeys.ToggleDisabledKey(control.Id, pressed);
	}

	public bool GetKeyState(IRibbonControl control)
	{
		return DisabledKeys.GetKeyState(control.Id);
	}

	public void PublishSharedSettings(IRibbonControl control)
	{
		SharedSettings.Publish(VH.A(169659));
	}

	public void Settings(IRibbonControl control)
	{
		if (EditMode.IsEditMode(MH.A.Application))
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (!Base.ConfigureMacabacus(MH.A.Application, (Microsoft.Office.Interop.PowerPoint.Application)null, (Microsoft.Office.Interop.Word.Application)null, KH.A, J.Ribbon))
			{
				return;
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				KH.A = new clsSettings();
				if (KH.A.UndoEnabled)
				{
					ExcelAddIn1.UndoRedo.Core.Enable();
				}
				else
				{
					ExcelAddIn1.UndoRedo.Core.Disable();
				}
				return;
			}
		}
	}

	public void SettingsBackup(IRibbonControl control)
	{
		clsSettings.SettingsExport();
	}

	public void SettingsRestore(IRibbonControl control)
	{
		clsSettings.SettingsImport();
	}

	public void SettingsReset(IRibbonControl control)
	{
		clsSettings.SettingsReset();
	}

	public bool ShowLMSCourseUrl(IRibbonControl control)
	{
		Profile userProfile = Base.UserProfile;
		object value;
		if (userProfile == null)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			value = null;
		}
		else
		{
			value = userProfile.LMSCourseUrl;
		}
		return !string.IsNullOrWhiteSpace((string)value);
	}

	public void HelpCenter(ref IRibbonControl control)
	{
		clsSupport.OnlineDocs(VH.A(205126));
	}

	public void EmailSupport(IRibbonControl control)
	{
		clsSupport.EmailSupport((CallingApp)1);
	}

	public void Feedback(IRibbonControl control)
	{
		Form.Show((OfficeApp)1);
	}

	public string GetSupportDescription(IRibbonControl control)
	{
		return clsSupport.GetSupportDescription();
	}

	public void GoToLMSCourse(IRibbonControl control)
	{
		clsSupport.GoToLMSCourse();
	}

	public void AboutMacabacus(ref IRibbonControl control)
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_000a: Expected O, but got Unknown
		UIFormsExtensions.CustomShowDialog((DependencyObject)new wpfAbout());
		_ = null;
	}

	public void Pronounce(ref IRibbonControl control)
	{
		clsUtilities.Pronounce();
	}

	public void Test(IRibbonControl control)
	{
	}

	public bool RepurposeToggle(IRibbonControl control, bool pressed, ref bool cancelDefault)
	{
		bool result = default(bool);
		return result;
	}

	public bool RepurposeButton(IRibbonControl control, ref bool cancelDefault)
	{
		bool result = default(bool);
		return result;
	}

	private bool A(IRibbonControl A, ref bool B)
	{
		try
		{
			if (BlockExecuteMso)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					B = false;
					break;
				}
			}
			else if (KH.A.UndoEnabled)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					RibbonXL.A(A.Id);
					B = true;
					break;
				}
			}
			else
			{
				B = false;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			B = false;
			ProjectData.ClearProjectError();
		}
		bool result = default(bool);
		return result;
	}

	private static void A(string A)
	{
		try
		{
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			Range a = (Range)application.Selection;
			bool flag = JH.A(a);
			BlockExecuteMso = true;
			application.CommandBars.ExecuteMso(A);
			uint num = TH.A(A);
			string b = default(string);
			if (num <= 1983181318)
			{
				if (num <= 1062001812)
				{
					while (true)
					{
						switch (5)
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
					if (num <= 155447713)
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
						if (num <= 80227853)
						{
							if (num != 59713405)
							{
								if (num != 80227853)
								{
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										break;
									}
								}
								else if (Operators.CompareString(A, VH.A(205290), TextCompare: false) == 0)
								{
									goto IL_0afa;
								}
							}
							else if (Operators.CompareString(A, VH.A(149846), TextCompare: false) != 0)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									break;
								}
							}
							else
							{
								b = VH.A(206355);
							}
						}
						else if (num != 106241546)
						{
							if (num != 139970389)
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									break;
								}
								if (num != 155447713)
								{
									while (true)
									{
										switch (6)
										{
										case 0:
											continue;
										}
										break;
									}
								}
								else if (Operators.CompareString(A, VH.A(205959), TextCompare: false) == 0)
								{
									goto IL_0b7f;
								}
							}
							else
							{
								if (Operators.CompareString(A, VH.A(206008), TextCompare: false) == 0)
								{
									goto IL_0b9d;
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									break;
								}
							}
						}
						else
						{
							if (Operators.CompareString(A, VH.A(206140), TextCompare: false) == 0)
							{
								goto IL_0b9d;
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								break;
							}
						}
					}
					else if (num <= 276554291)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							break;
						}
						if (num != 229132862)
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
							if (num != 276554291)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									break;
								}
							}
							else if (Operators.CompareString(A, VH.A(205330), TextCompare: false) == 0)
							{
								b = VH.A(148035);
							}
						}
						else
						{
							if (Operators.CompareString(A, VH.A(205467), TextCompare: false) == 0)
							{
								goto IL_0b70;
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								break;
							}
						}
					}
					else if (num != 429279826)
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
						if (num != 996939909)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								break;
							}
							if (num != 1062001812)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									break;
								}
							}
							else if (Operators.CompareString(A, VH.A(205353), TextCompare: false) != 0)
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
							}
							else
							{
								b = VH.A(206421);
							}
						}
						else if (Operators.CompareString(A, VH.A(205586), TextCompare: false) == 0)
						{
							goto IL_0b7f;
						}
					}
					else if (Operators.CompareString(A, VH.A(169867), TextCompare: false) != 0)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							break;
						}
					}
					else
					{
						b = VH.A(148976);
					}
				}
				else if (num <= 1367952463)
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
					if (num <= 1185225039)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							break;
						}
						if (num != 1165397398)
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
							if (num != 1185225039)
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
							}
							else
							{
								if (Operators.CompareString(A, VH.A(205265), TextCompare: false) == 0)
								{
									goto IL_0afa;
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									break;
								}
							}
						}
						else if (Operators.CompareString(A, VH.A(205179), TextCompare: false) == 0)
						{
							b = VH.A(205179);
						}
					}
					else if (num != 1190436747)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							break;
						}
						if (num != 1273349792)
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
							if (num == 1367952463)
							{
								if (Operators.CompareString(A, VH.A(205232), TextCompare: false) == 0)
								{
									goto IL_0afa;
								}
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									break;
								}
							}
						}
						else
						{
							if (Operators.CompareString(A, VH.A(205875), TextCompare: false) == 0)
							{
								goto IL_0b7f;
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								break;
							}
						}
					}
					else if (Operators.CompareString(A, VH.A(205311), TextCompare: false) != 0)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							break;
						}
					}
					else
					{
						b = VH.A(206392);
					}
				}
				else if (num <= 1718516818)
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
					if (num != 1706366601)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							break;
						}
						if (num == 1718516818)
						{
							if (Operators.CompareString(A, VH.A(205201), TextCompare: false) != 0)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									break;
								}
							}
							else
							{
								b = VH.A(206285);
							}
						}
					}
					else
					{
						if (Operators.CompareString(A, VH.A(205803), TextCompare: false) == 0)
						{
							goto IL_0b7f;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							break;
						}
					}
				}
				else if (num != 1810622612)
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
					if (num != 1926406787)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							break;
						}
						if (num != 1983181318)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								break;
							}
						}
						else if (Operators.CompareString(A, VH.A(205401), TextCompare: false) != 0)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								break;
							}
						}
						else
						{
							b = VH.A(206479);
						}
					}
					else
					{
						if (Operators.CompareString(A, VH.A(205766), TextCompare: false) == 0)
						{
							goto IL_0b7f;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							break;
						}
					}
				}
				else if (Operators.CompareString(A, VH.A(205374), TextCompare: false) != 0)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						break;
					}
				}
				else
				{
					b = VH.A(206452);
				}
			}
			else if (num <= 3246123585u)
			{
				if (num <= 2592407839u)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					if (num <= 2485873578u)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							break;
						}
						if (num != 2377171085u)
						{
							if (num != 2485873578u)
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
							}
							else if (Operators.CompareString(A, VH.A(206187), TextCompare: false) == 0)
							{
								goto IL_0b9d;
							}
						}
						else if (Operators.CompareString(A, VH.A(205188), TextCompare: false) != 0)
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
						}
						else
						{
							b = VH.A(205188);
						}
					}
					else if (num != 2508154221u)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							break;
						}
						if (num != 2564959815u)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								break;
							}
							if (num != 2592407839u)
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
							}
							else if (Operators.CompareString(A, VH.A(151804), TextCompare: false) != 0)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									break;
								}
							}
							else
							{
								b = VH.A(151804);
							}
						}
						else
						{
							if (Operators.CompareString(A, VH.A(205912), TextCompare: false) == 0)
							{
								goto IL_0b7f;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								break;
							}
						}
					}
					else if (Operators.CompareString(A, VH.A(205621), TextCompare: false) == 0)
					{
						goto IL_0b7f;
					}
				}
				else if (num <= 2899771486u)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					if (num != 2695166007u)
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
						if (num != 2899771486u)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								break;
							}
						}
						else if (Operators.CompareString(A, VH.A(150594), TextCompare: false) != 0)
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
						}
						else
						{
							b = VH.A(150611);
						}
					}
					else
					{
						if (Operators.CompareString(A, VH.A(206081), TextCompare: false) == 0)
						{
							goto IL_0b9d;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							break;
						}
					}
				}
				else if (num != 2940351310u)
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
					if (num != 3113897897u)
					{
						if (num != 3246123585u)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								break;
							}
						}
						else if (Operators.CompareString(A, VH.A(149813), TextCompare: false) != 0)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								break;
							}
						}
						else
						{
							b = VH.A(206318);
						}
					}
					else
					{
						if (Operators.CompareString(A, VH.A(205697), TextCompare: false) == 0)
						{
							goto IL_0b7f;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							break;
						}
					}
				}
				else
				{
					if (Operators.CompareString(A, VH.A(205506), TextCompare: false) == 0)
					{
						goto IL_0b70;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						break;
					}
				}
			}
			else if (num <= 3516201001u)
			{
				if (num <= 3315270155u)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						break;
					}
					if (num != 3273491695u)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							break;
						}
						if (num != 3315270155u)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								break;
							}
						}
						else
						{
							if (Operators.CompareString(A, VH.A(206234), TextCompare: false) == 0)
							{
								goto IL_0b9d;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								break;
							}
						}
					}
					else
					{
						if (Operators.CompareString(A, VH.A(205840), TextCompare: false) == 0)
						{
							goto IL_0b7f;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							break;
						}
					}
				}
				else if (num != 3369241730u)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						break;
					}
					if (num != 3501011684u)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							break;
						}
						if (num != 3516201001u)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								break;
							}
						}
						else
						{
							if (Operators.CompareString(A, VH.A(149451), TextCompare: false) == 0)
							{
								goto IL_0b70;
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								break;
							}
						}
					}
					else
					{
						if (Operators.CompareString(A, VH.A(205658), TextCompare: false) == 0)
						{
							goto IL_0b7f;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							break;
						}
					}
				}
				else
				{
					if (Operators.CompareString(A, VH.A(205739), TextCompare: false) == 0)
					{
						goto IL_0b7f;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						break;
					}
				}
			}
			else if (num <= 3717847479u)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				if (num != 3677912017u)
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
					if (num != 3717847479u)
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
					}
					else if (Operators.CompareString(A, VH.A(205718), TextCompare: false) == 0)
					{
						goto IL_0b7f;
					}
				}
				else if (Operators.CompareString(A, VH.A(170065), TextCompare: false) != 0)
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
				}
				else
				{
					b = VH.A(148997);
				}
			}
			else if (num != 3882872155u)
			{
				if (num != 3883787349u)
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
					if (num == 4252145872u)
					{
						if (Operators.CompareString(A, VH.A(205434), TextCompare: false) != 0)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								break;
							}
						}
						else
						{
							b = VH.A(206534);
						}
					}
				}
				else if (Operators.CompareString(A, VH.A(149418), TextCompare: false) == 0)
				{
					goto IL_0b70;
				}
			}
			else if (Operators.CompareString(A, VH.A(205545), TextCompare: false) == 0)
			{
				goto IL_0b7f;
			}
			goto IL_0bc8;
			IL_0b7f:
			b = VH.A(146542);
			goto IL_0bc8;
			IL_0afa:
			b = VH.A(187074);
			goto IL_0bc8;
			IL_0b9d:
			b = VH.A(196819);
			goto IL_0bc8;
			IL_0b70:
			b = VH.A(148068);
			goto IL_0bc8;
			IL_0bc8:
			_ = null;
			if (flag)
			{
				JH.A(a, b);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		finally
		{
			BlockExecuteMso = false;
			Range a = null;
		}
	}

	public void RepurposeCopyCut(IRibbonControl control, ref bool cancelDefault)
	{
		if (SoftDisable.IsEnabled)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					try
					{
						string id = control.Id;
						if (Operators.CompareString(id, VH.A(224), TextCompare: false) != 0)
						{
							if (Operators.CompareString(id, VH.A(197247), TextCompare: false) != 0)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										break;
									default:
										return;
									}
								}
							}
							Paste.Cut();
						}
						else
						{
							Paste.Copy();
						}
						return;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
						return;
					}
					finally
					{
						cancelDefault = false;
					}
				}
			}
		}
		cancelDefault = false;
	}

	public void RepurposeSheetUnhide(IRibbonControl control, ref bool cancelDefault)
	{
		if (SoftDisable.IsEnabled)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					try
					{
						if (!Workbooks.IsShared(MH.A.Application.ActiveWorkbook, false, (System.Windows.Window)null))
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									ExcelAddIn1.Sheets.Visibility.D();
									cancelDefault = true;
									return;
								}
							}
						}
						cancelDefault = false;
						return;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						cancelDefault = false;
						ProjectData.ClearProjectError();
						return;
					}
				}
			}
		}
		cancelDefault = false;
	}

	public void UnhideSheets(ref IRibbonControl control)
	{
		if (Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
		{
			ExcelAddIn1.Sheets.Visibility.D();
		}
	}
}
