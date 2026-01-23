using System;
using System.Drawing;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Windows;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.Config;
using MacabacusMacros.Feedback;
using MacabacusMacros.ImportExport;
using MacabacusMacros.Libraries.Manage;
using MacabacusMacros.UI;
using MacabacusMacros.UI.FormsExtensions;
using Macabacus_Word.Aiwa;
using Macabacus_Word.Colors;
using Macabacus_Word.DocBuilder;
using Macabacus_Word.Keyboard;
using Macabacus_Word.Library2;
using Macabacus_Word.Library2.UI;
using Macabacus_Word.Library2.Versioning;
using Macabacus_Word.Links;
using Macabacus_Word.Proofing;
using Macabacus_Word.Proofing.UI;
using Macabacus_Word.Publishing;
using Macabacus_Word.Shapes;
using Macabacus_Word.TextOps;
using Macabacus_Word.TextOps.Redaction;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;
using stdole;

namespace Macabacus_Word;

[ComVisible(true)]
public sealed class RibbonWD : IRibbonExtensibility
{
	private IRibbonUI m_A;

	public string GetCustomUI(string ribbonID)
	{
		string outerXml = default(string);
		try
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(M.Ribbon);
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
		NC.A = ribbonUI;
		clsRibbon.Ribbon = ribbonUI;
		Licensing.Authenticate();
	}

	public IPictureDisp CallbackLoadImage(string resourceName)
	{
		return OC.A((Bitmap)new ResourceManager(XC.A(341), Assembly.GetExecutingAssembly()).GetObject(resourceName));
	}

	public string MacabacusKeyTip(IRibbonControl control)
	{
		return clsRibbon.MacabacusTabKeyTip(XC.A(18421));
	}

	public void ActivateMacabacus(ref IRibbonControl control)
	{
		Licensing.Activate();
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

	public bool DocumentOpen_Callback(IRibbonControl control)
	{
		bool result;
		try
		{
			result = PC.A.Application.Documents.Count != 0;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public bool IsLinkedItem_Callback(IRibbonControl control)
	{
		return A();
	}

	private bool A()
	{
		return Common.IsLinkSelected();
	}

	private bool B()
	{
		bool result;
		try
		{
			Microsoft.Office.Interop.Word.Selection selection = PC.A.Application.ActiveWindow.Selection;
			result = (selection.Type == WdSelectionType.wdSelectionShape) | (selection.Type == WdSelectionType.wdSelectionInlineShape);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public bool IsShapeSelected_Callback(IRibbonControl control)
	{
		return B();
	}

	public bool CanImport_Callback(IRibbonControl control)
	{
		return true;
	}

	public string GetScreenTip(IRibbonControl control)
	{
		return Shortcuts.HotkeyScreenTip(control.Id);
	}

	public string GetSuperTip(IRibbonControl control)
	{
		return Shortcuts.SuperTip(control.Id);
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
		return clsColors.FontColorMenu((ColorRole[])(object)new ColorRole[1]);
	}

	public string FillColorMenu(IRibbonControl control)
	{
		return clsColors.FillColorMenu((ColorRole[])(object)new ColorRole[1]);
	}

	public string BorderColorMenu(IRibbonControl control)
	{
		return clsColors.BorderColorMenu((ColorRole[])(object)new ColorRole[1]);
	}

	public void DoFontColor(IRibbonControl control)
	{
		Macabacus_Word.Colors.Font.A(control.Tag);
	}

	public void DoFillColor(IRibbonControl control)
	{
		Fill.A(control.Tag);
	}

	public void DoBorderColor(IRibbonControl control)
	{
		Macabacus_Word.Colors.Border.A(control.Tag);
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
		Macabacus_Word.Colors.Font.A(index);
	}

	public void FillColorAction(IRibbonControl control, string id, int index)
	{
		Fill.A(index);
	}

	public void BorderColorAction(IRibbonControl control, string id, int index)
	{
		Macabacus_Word.Colors.Border.A(index);
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
		Macabacus_Word.Colors.Font.ButtonColor();
	}

	public void FillColorButton(IRibbonControl control)
	{
		Fill.ButtonColor();
	}

	public void BorderColorButton(IRibbonControl control)
	{
		Macabacus_Word.Colors.Border.ButtonColor();
	}

	public void BorderNone(IRibbonControl control)
	{
		Macabacus_Word.Colors.Border.NoBorder();
	}

	public void NoFill(IRibbonControl control)
	{
		Fill.NoFill();
	}

	public void CycleFontColor(IRibbonControl control)
	{
		Macabacus_Word.Colors.Font.Cycle();
	}

	public void CycleFillColor(IRibbonControl control)
	{
		Fill.Cycle();
	}

	public void CycleBorderColor(IRibbonControl control)
	{
		Macabacus_Word.Colors.Border.Cycle();
	}

	public void InsertShape(ref IRibbonControl control, bool pressed)
	{
		Macabacus_Word.Library2.UI.Pane.Toggle(pressed);
	}

	public bool IsContentPaneOpen(IRibbonControl control)
	{
		return Macabacus_Word.Library2.UI.Pane.IsVisible();
	}

	public void InsertSymbol(IRibbonControl control)
	{
		Symbols.Insert(Conversions.ToInteger(control.Tag));
	}

	public void RedactSelection(IRibbonControl control)
	{
		Redact.RedactSelection();
	}

	public void FindAndRedact(IRibbonControl control)
	{
		Redact.FindAndRedact();
	}

	public void NumToWordsPlain(IRibbonControl control)
	{
		Numbers.NumberToPlainWords();
	}

	public void NumToWordsLegal(IRibbonControl control)
	{
		Numbers.NumberToLegalWords();
	}

	public void ApplyHeadingStyle(IRibbonControl control)
	{
		clsStyles.ApplyHeadingStyle(control.Tag);
	}

	public void ApplyListStyle(IRibbonControl control)
	{
		clsStyles.ApplyListStyle(control.Tag);
	}

	public void ApplyTextStyle(IRibbonControl control)
	{
		clsStyles.ApplyTextStyle(control.Tag);
	}

	public void ApplyTableStyle(IRibbonControl control)
	{
		clsStyles.ApplyTableStyle(control.Tag);
	}

	public string MenuHeadingStyles(IRibbonControl control)
	{
		return clsStyles.MenuHeadingStyles();
	}

	public string MenuBulletStyles(IRibbonControl control)
	{
		return clsStyles.MenuListStyles();
	}

	public string MenuTextStyles(IRibbonControl control)
	{
		return clsStyles.MenuTextStyles();
	}

	public string MenuTableStyles(IRibbonControl control)
	{
		return clsStyles.MenuTableStyles();
	}

	public void ImportExcel(IRibbonControl control)
	{
		Macabacus_Word.Links.ImportExcel.A();
	}

	public void ImportExcelAsGraphic(IRibbonControl control)
	{
		Macabacus_Word.Links.ImportExcel.B();
	}

	public void ImportExcelAsImage(IRibbonControl control)
	{
		Macabacus_Word.Links.ImportExcel.C();
	}

	public void ImportExcelAsTable(IRibbonControl control)
	{
		Macabacus_Word.Links.ImportExcel.D();
	}

	public void ImportExcelAsEmbedded(IRibbonControl control)
	{
		Macabacus_Word.Links.ImportExcel.E();
	}

	public void ImportExcelAsText(IRibbonControl control)
	{
		Macabacus_Word.Links.ImportExcel.F();
	}

	public void ImportExcelAsChart(IRibbonControl control)
	{
		Macabacus_Word.Links.ImportExcel.G();
	}

	public void PrintAreasSheet(IRibbonControl control)
	{
		PagesToWord.PrintAreasSelectedSheets((Microsoft.Office.Interop.Excel.Application)null, PC.A.Application);
	}

	public void PrintAreasAll(IRibbonControl control)
	{
		PagesToWord.PrintAreasAllSheets((Microsoft.Office.Interop.Excel.Application)null, PC.A.Application);
	}

	public void MatchWidth(ref IRibbonControl control, bool pressed)
	{
		N.Settings.ImportMatchDestinationWidth = pressed;
		if (!pressed)
		{
			return;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			N.Settings.ImportMatchDestinationHeight = false;
			NC.A.InvalidateControl(XC.A(42590));
			return;
		}
	}

	public void MatchHeight(ref IRibbonControl control, bool pressed)
	{
		N.Settings.ImportMatchDestinationHeight = pressed;
		if (!pressed)
		{
			return;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			N.Settings.ImportMatchDestinationWidth = false;
			NC.A.InvalidateControl(XC.A(42613));
			return;
		}
	}

	public bool GetMatchWidthChecked(IRibbonControl control)
	{
		return N.Settings.ImportMatchDestinationWidth;
	}

	public bool GetMatchHeightChecked(IRibbonControl control)
	{
		return N.Settings.ImportMatchDestinationHeight;
	}

	public void LinkedWizard(IRibbonControl control)
	{
		Manage.LinkWizard();
	}

	public void EditSource(IRibbonControl control)
	{
		Edit.EditLink();
	}

	public void ViewSource(IRibbonControl control)
	{
		Macabacus_Word.Links.View.ViewSource();
	}

	public void UpdateLink(IRibbonControl control)
	{
		Refresh.SelectedLinks();
	}

	public void UpdateAllLinks(IRibbonControl control)
	{
		Refresh.UpdateAllLinks();
	}

	public void RemoveLink(IRibbonControl control)
	{
		Break.SelectedLinks();
	}

	public void HighlightLinks(IRibbonControl control)
	{
		Highlight.Add();
	}

	public void UnhighlightLinks(IRibbonControl control)
	{
		Highlight.Remove();
	}

	public void HideTextLinks(IRibbonControl control)
	{
		Macabacus_Word.Links.Visibility.HideTextLinks();
	}

	public void ShowTextLinks(IRibbonControl control)
	{
		Macabacus_Word.Links.Visibility.ShowTextLinks();
	}

	public void AiwaToggle(ref IRibbonControl control, bool pressed)
	{
		Macabacus_Word.Aiwa.Pane.Toggle(pressed);
	}

	public bool IsAiwaOpen(IRibbonControl control)
	{
		return Macabacus_Word.Aiwa.Pane.IsVisible();
	}

	public string DocumentLibraryMenu(IRibbonControl control)
	{
		return Macabacus_Word.Library2.Documents.BuildFilesMenu();
	}

	public void OpenFile(IRibbonControl control)
	{
		Macabacus_Word.Library2.Documents.OpenFile(control.Tag);
	}

	public void LibraryShowAll(IRibbonControl control)
	{
		Macabacus_Word.Library2.UI.Pane.Show(blnShapes: true, blnImages: true, blnCharts: true, blnText: true, blnDocs: true, blnPDFs: true);
	}

	public void LibraryShowShapes(IRibbonControl control)
	{
		Macabacus_Word.Library2.UI.Pane.Show(blnShapes: true, blnImages: false, blnCharts: false, blnText: false);
	}

	public void LibraryShowImages(IRibbonControl control)
	{
		Macabacus_Word.Library2.UI.Pane.Show(blnShapes: false, blnImages: true, blnCharts: false, blnText: false);
	}

	public void LibraryShowCharts(IRibbonControl control)
	{
		Macabacus_Word.Library2.UI.Pane.Show(blnShapes: false, blnImages: false, blnCharts: true, blnText: false);
	}

	public void LibraryShowText(IRibbonControl control)
	{
		Macabacus_Word.Library2.UI.Pane.Show(blnShapes: false, blnImages: false, blnCharts: false, blnText: true);
	}

	public void ManageLibraryContents(IRibbonControl control)
	{
		Admin.LibraryManager((Microsoft.Office.Interop.PowerPoint.Application)null, (Microsoft.Office.Interop.Excel.Application)null, PC.A.Application);
	}

	public void LibraryVersionControl(IRibbonControl control, bool pressed)
	{
		Macabacus_Word.Library2.Versioning.Pane.Toggle(pressed);
	}

	public bool IsLibContentPaneOpen(IRibbonControl control)
	{
		return Macabacus_Word.Library2.Versioning.Pane.IsVisible();
	}

	public void HelpCenter(IRibbonControl control)
	{
		clsSupport.OnlineDocs(XC.A(42634));
	}

	public void EmailSupport(IRibbonControl control)
	{
		clsSupport.EmailSupport((CallingApp)3);
	}

	public void Feedback(IRibbonControl control)
	{
		Form.Show((OfficeApp)3);
	}

	public string GetSupportDescription(IRibbonControl control)
	{
		return clsSupport.GetSupportDescription();
	}

	public void AboutMacabacus(IRibbonControl control)
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_000a: Expected O, but got Unknown
		UIFormsExtensions.CustomShowDialog((DependencyObject)new wpfAbout());
		_ = null;
	}

	public void Pronounce(IRibbonControl control)
	{
		clsUtilities.Pronounce();
	}

	public void ProofingToggle(ref IRibbonControl control, bool pressed)
	{
		Macabacus_Word.Proofing.UI.Pane.Toggle(pressed);
	}

	public bool IsProofingPaneOpen(IRibbonControl control)
	{
		return Macabacus_Word.Proofing.UI.Pane.IsVisible();
	}

	public void ProofDocument(IRibbonControl control)
	{
		Macabacus_Word.Proofing.UI.Pane.CheckDocument();
	}

	public void ProofSelection(IRibbonControl control)
	{
		Macabacus_Word.Proofing.UI.Pane.CheckSelection();
	}

	public string ProofingLanguageMenu(IRibbonControl control)
	{
		return Language.LanguagesMenu();
	}

	public void SetLanguage(IRibbonControl control)
	{
		Language.SetProofingLanguage(control);
	}

	public void EmailDocument(IRibbonControl control)
	{
		Send.ShowDialog();
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

	public void SaveAll(ref IRibbonControl control)
	{
		clsFile.SaveAll();
	}

	public void SaveUp(IRibbonControl control)
	{
		clsFile.SaveUp();
	}

	public void CloseOthers(IRibbonControl control)
	{
		clsFile.CloseOthers();
	}

	public void Reopen(IRibbonControl control)
	{
		clsFile.Reopen();
	}

	public void Duplicate(IRibbonControl control)
	{
		clsFile.Duplicate();
	}

	public void ShowInFolder(IRibbonControl control)
	{
		clsFile.OpenFolder();
	}

	public void CopyPath(IRibbonControl control)
	{
		clsFile.CopyPath();
	}

	public void PdfToFolder(ref IRibbonControl control)
	{
		Pdf.ToFolder();
	}

	public void PublishSharedSettings(IRibbonControl control)
	{
		SharedSettings.Publish(XC.A(18421));
	}

	public void Settings(IRibbonControl control)
	{
		if (!Base.ConfigureMacabacus((Microsoft.Office.Interop.Excel.Application)null, (Microsoft.Office.Interop.PowerPoint.Application)null, PC.A.Application, NC.A, M.Ribbon))
		{
			return;
		}
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
			NC.A = new clsSettings();
			return;
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
		Shortcuts.OverrideHotkeys();
	}

	public void PrintHotkeys(IRibbonControl control)
	{
	}

	public void ImportHotkeys(IRibbonControl control)
	{
		Shortcuts.ImportHotkeys();
	}

	public void ToggleDisabledKey(ref IRibbonControl control, bool pressed)
	{
		DisabledKeys.ToggleDisabledKey(control.Id, pressed);
	}

	public bool GetKeyState(IRibbonControl control)
	{
		return DisabledKeys.GetKeyState(control.Id);
	}

	public bool IsDocBuilderOpen(IRibbonControl control)
	{
		return Macabacus_Word.DocBuilder.Pane.IsVisible();
	}

	public void DocBuilderToggle(ref IRibbonControl control, bool pressed)
	{
		Macabacus_Word.DocBuilder.Pane.Toggle(pressed);
	}

	public bool ShowBetaTools(IRibbonControl control)
	{
		return clsRibbon.ShowBetaTools;
	}

	public bool ShowNewerVersionNotice(IRibbonControl control)
	{
		return clsUpdate.ShowNewerVersionNotice(NC.A);
	}

	public void DownloadUpdate(ref IRibbonControl control)
	{
		clsUpdate.DownloadUpdate(NC.A);
	}

	public void DismissUpdate(ref IRibbonControl control)
	{
		clsUpdate.DismissUpdate(NC.A);
	}

	public string UpdateLabel(IRibbonControl control)
	{
		return clsUpdate.NewerVersionLabel();
	}

	public void BackstageShow(object contextObject)
	{
		NC.B = true;
	}

	public void BackstageHide(object contextObject)
	{
		NC.B = false;
	}
}
