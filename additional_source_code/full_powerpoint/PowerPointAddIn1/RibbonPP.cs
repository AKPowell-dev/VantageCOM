using System;
using System.Collections;
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
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Agenda;
using PowerPointAddIn1.Aiwa;
using PowerPointAddIn1.Colors;
using PowerPointAddIn1.Colors.Recolor;
using PowerPointAddIn1.DeckCheck;
using PowerPointAddIn1.DeckCheck.UI;
using PowerPointAddIn1.Explorer;
using PowerPointAddIn1.FormatPainter;
using PowerPointAddIn1.Library2;
using PowerPointAddIn1.Library2.UI;
using PowerPointAddIn1.Library2.Versioning;
using PowerPointAddIn1.Links;
using PowerPointAddIn1.LogoLibrary;
using PowerPointAddIn1.MasterShapes;
using PowerPointAddIn1.Pagination;
using PowerPointAddIn1.Presentation;
using PowerPointAddIn1.Publishing;
using PowerPointAddIn1.Publishing.Share;
using PowerPointAddIn1.Shapes;
using PowerPointAddIn1.Shapes.Arrange;
using PowerPointAddIn1.Shapes.SelectMatch;
using PowerPointAddIn1.Shapes.Templated;
using PowerPointAddIn1.Slides;
using PowerPointAddIn1.Template;
using PowerPointAddIn1.Template.Wizard;
using PowerPointAddIn1.TextOps;
using PowerPointAddIn1.TurboShapes;
using stdole;

namespace PowerPointAddIn1;

[ComVisible(true)]
public sealed class RibbonPP : IRibbonExtensibility
{
	private IRibbonUI m_A;

	public string GetCustomUI(string ribbonID)
	{
		string outerXml = default(string);
		try
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(OB.Ribbon);
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
		KG.A = ribbonUI;
		clsRibbon.Ribbon = ribbonUI;
		Licensing.Authenticate();
	}

	public IPictureDisp CallbackLoadImage(string resourceName)
	{
		return MG.A((Bitmap)new ResourceManager(AH.A(1885), Assembly.GetExecutingAssembly()).GetObject(resourceName));
	}

	public string MacabacusKeyTip(IRibbonControl control)
	{
		return clsRibbon.MacabacusTabKeyTip(AH.A(116727));
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

	public void Test(IRibbonControl control)
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = application.Presentations.Add();
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(58943)).RemoveEventHandler(application, new EApplication_PresentationNewSlideEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(58943)).AddEventHandler(application, new EApplication_PresentationNewSlideEventHandler(A));
		presentation.Slides.AddSlide(1, presentation.SlideMaster.CustomLayouts[1]);
		presentation = null;
		application = null;
	}

	private void A(Slide A)
	{
		Interaction.MsgBox(A.SlideID);
	}

	public void InsertTitle(IRibbonControl control)
	{
		Create.InsertTitleSlide();
	}

	public bool ShowTocMenu(IRibbonControl control)
	{
		return Flysheets.ShowTocMenu();
	}

	public void InsertTableOfContents(IRibbonControl control)
	{
		Create.InsertTocSlide();
	}

	public void InsertFlysheet(IRibbonControl control)
	{
		Flysheets.InsertFlysheet();
	}

	public void InsertLegal(IRibbonControl control)
	{
		Create.InsertLegalSlide();
	}

	public void InsertContact(IRibbonControl control)
	{
		Create.InsertContactSlide();
	}

	public void InsertBlank(IRibbonControl control)
	{
		Create.InsertBlankSlide();
	}

	public void InsertFrontCover(IRibbonControl control)
	{
		Create.InsertFrontCoverSlide();
	}

	public void InsertBackCover(IRibbonControl control)
	{
		Create.InsertBackCoverSlide();
	}

	public void InsertContent(IRibbonControl control)
	{
		InsertSlide.ShowDialog();
	}

	public void PaginateToggle(ref IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.Pagination.Pane.Toggle(pressed);
	}

	public bool IsPaginateOpen(IRibbonControl control)
	{
		return PowerPointAddIn1.Pagination.Pane.IsVisible();
	}

	public void ResetSlideNumbers(IRibbonControl control)
	{
		PowerPointAddIn1.Pagination.SlideNumbers.Reset();
	}

	public void MarkFacingSlide(IRibbonControl control)
	{
		FacingSlides.MarkSlide();
	}

	public void UnmarkFacingSlide(IRibbonControl control)
	{
		FacingSlides.UnmarkSlide();
	}

	public void UpdateTableOfContents(IRibbonControl control)
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		application.StartNewUndoEntry();
		Update.A(A: false, application.ActivePresentation);
		application = null;
	}

	public void ShowInTableOfContents(IRibbonControl control)
	{
		TableOfContents.A(Conversions.ToInteger(control.Tag));
	}

	public void ExcludeFromTableOfContents(IRibbonControl control)
	{
		TableOfContents.A();
	}

	public void ClearTableOfContentsInputs(IRibbonControl control)
	{
		TableOfContents.B();
	}

	public void FlysheetLevel(IRibbonControl control)
	{
		Flysheets.FlysheetLevel(Conversions.ToInteger(control.Tag));
	}

	public void FlysheetStyleTopic(ref IRibbonControl control, bool pressed)
	{
		Flysheets.FlysheetStyleTopic();
	}

	public void FlysheetStyleAgenda(ref IRibbonControl control, bool pressed)
	{
		Flysheets.FlysheetStyleAgenda();
	}

	public void FlysheetCollapse(ref IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.Agenda.Behavior.ToggleAutoCollapse(pressed);
	}

	public void FlysheetSkip(ref IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.Agenda.Behavior.ToggleSkipDoubles(pressed);
	}

	public void FlysheetOmit(ref IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.Agenda.Behavior.ToggleOmitDoubles(pressed);
	}

	public void TocShowSubsections(ref IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.Agenda.Behavior.ToggleShowSubsections(pressed);
	}

	public bool FlysheetBehaviorEnabled(IRibbonControl control)
	{
		return PowerPointAddIn1.Agenda.Behavior.FlysheetBehaviorEnabled();
	}

	public bool FlysheetTopicPressed(IRibbonControl control)
	{
		return PowerPointAddIn1.Agenda.Behavior.FlysheetTopicPressed();
	}

	public bool FlysheetAgendaPressed(IRibbonControl control)
	{
		return PowerPointAddIn1.Agenda.Behavior.FlysheetAgendaPressed();
	}

	public bool FlysheetCollapsePressed(IRibbonControl control)
	{
		return PowerPointAddIn1.Agenda.Behavior.FlysheetCollapsePressed();
	}

	public bool FlysheetSkipPressed(IRibbonControl control)
	{
		return PowerPointAddIn1.Agenda.Behavior.FlysheetSkipPressed();
	}

	public bool FlysheetOmitPressed(IRibbonControl control)
	{
		return PowerPointAddIn1.Agenda.Behavior.FlysheetOmitPressed();
	}

	public bool ShowSubsectionsPressed(IRibbonControl control)
	{
		return PowerPointAddIn1.Agenda.Behavior.ShowSubsectionsPressed();
	}

	public void FileNew_Repurposed(IRibbonControl control, ref bool cancelDefault)
	{
		Create.FileNew_Repurposed(ref cancelDefault);
	}

	public void SectionAdd_Repurposed(IRibbonControl control, ref bool cancelDefault)
	{
		cancelDefault = Sections.A();
	}

	public void SectionRename_Repurposed(IRibbonControl control, ref bool cancelDefault)
	{
		cancelDefault = Sections.B();
	}

	public void SectionMergeWithPrevious_Repurposed(IRibbonControl control, ref bool cancelDefault)
	{
		cancelDefault = Sections.C();
	}

	public void SectionDelete_Repurposed(IRibbonControl control, ref bool cancelDefault)
	{
		cancelDefault = Sections.D();
	}

	public void SectionRemoveAll_Repurposed(IRibbonControl control, ref bool cancelDefault)
	{
		Sections.A();
		cancelDefault = false;
	}

	public void SectionMoveUp_Repurposed(IRibbonControl control, ref bool cancelDefault)
	{
		cancelDefault = Sections.E();
	}

	public void SectionMoveDown_Repurposed(IRibbonControl control, ref bool cancelDefault)
	{
		cancelDefault = Sections.F();
	}

	public void SectionPromote(IRibbonControl control)
	{
		Sections.B();
	}

	public void SectionDemote(IRibbonControl control)
	{
		Sections.C();
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
		PowerPointAddIn1.Colors.Ribbon.A(control.Tag);
	}

	public void DoFillColor(IRibbonControl control)
	{
		PowerPointAddIn1.Colors.Ribbon.B(control.Tag);
	}

	public void DoBorderColor(IRibbonControl control)
	{
		PowerPointAddIn1.Colors.Ribbon.C(control.Tag);
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
		PowerPointAddIn1.Colors.Ribbon.A(index);
	}

	public void FillColorAction(IRibbonControl control, string id, int index)
	{
		PowerPointAddIn1.Colors.Ribbon.B(index);
	}

	public void BorderColorAction(IRibbonControl control, string id, int index)
	{
		PowerPointAddIn1.Colors.Ribbon.C(index);
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
		PowerPointAddIn1.Colors.Ribbon.A();
	}

	public void FillColorButton(IRibbonControl control)
	{
		PowerPointAddIn1.Colors.Ribbon.B();
	}

	public void BorderColorButton(IRibbonControl control)
	{
		PowerPointAddIn1.Colors.Ribbon.D();
	}

	public void BorderNone(IRibbonControl control)
	{
		PowerPointAddIn1.Colors.Ribbon.E();
	}

	public void NoFill(IRibbonControl control)
	{
		PowerPointAddIn1.Colors.Ribbon.C();
	}

	public void MakeOpaque(IRibbonControl control)
	{
		FillTransparency.Fix();
	}

	public void FreezePresentation(IRibbonControl control)
	{
		Freeze.A(A: false);
	}

	public void FreezeSelection(IRibbonControl control)
	{
		Freeze.A();
	}

	public void HarveyBall(IRibbonControl control)
	{
		PowerPointAddIn1.TurboShapes.HarveyBall.Add();
	}

	public void ProgressBar(IRibbonControl control)
	{
		PowerPointAddIn1.TurboShapes.ProgressBar.Add();
	}

	public void RatingBar(IRibbonControl control)
	{
		PowerPointAddIn1.TurboShapes.RatingBar.Add();
	}

	public void Arrow(IRibbonControl control)
	{
		PowerPointAddIn1.TurboShapes.Arrow.Add();
	}

	public void CheckBox(IRibbonControl control)
	{
		PowerPointAddIn1.TurboShapes.CheckBox.Add();
	}

	public void NoticeIcon(IRibbonControl control)
	{
		PowerPointAddIn1.TurboShapes.NoticeIcon.Add();
	}

	public void TrafficLight(IRibbonControl control)
	{
		PowerPointAddIn1.TurboShapes.TrafficLight.Add();
	}

	public void Tachometer(IRibbonControl control)
	{
		PowerPointAddIn1.TurboShapes.Tachometer.Add();
	}

	public void Thermometer(IRibbonControl control)
	{
		PowerPointAddIn1.TurboShapes.Thermometer.Add();
	}

	public void ToggleSwitch(IRibbonControl control)
	{
		PowerPointAddIn1.TurboShapes.ToggleSwitch.Add();
	}

	public void SliderBar(IRibbonControl control)
	{
		PowerPointAddIn1.TurboShapes.SliderBar.Add();
	}

	public void CreateTurboShape(IRibbonControl control)
	{
		Custom.Convert();
	}

	public void ToggleFormatPainter(ref IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.FormatPainter.Pane.Toggle(pressed);
	}

	public bool IsFormatPainterOpen(IRibbonControl control)
	{
		return PowerPointAddIn1.FormatPainter.Pane.IsVisible();
	}

	public void CopyProperties(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.Copy();
	}

	public void ApplySize(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.Size();
	}

	public void ApplyHeight(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.Height();
	}

	public void ApplyWidth(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.Width();
	}

	public void ApplyPosition(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.Position();
	}

	public void ApplyTop(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.Top();
	}

	public void ApplyLeft(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.Left();
	}

	public void ApplyMidpointY(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.MidpointY();
	}

	public void ApplyMidpointX(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.MidpointX();
	}

	public void ApplyRotation(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.Rotation();
	}

	public void ApplyLockAspect(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.LockAspectRatio();
	}

	public void ApplyBullets(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.Bullets();
	}

	public void ApplyIndents(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.Indents();
	}

	public void ApplyLineSpacing(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.LineSpacing();
	}

	public void ApplyMargins(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.Margins();
	}

	public void ApplyTextWrap(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.TextWrap();
	}

	public void ApplyAutoSize(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.AutoSize();
	}

	public void ApplyAlignH(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.HorizontalAlignment();
	}

	public void ApplyAlignV(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.VerticalAlignment();
	}

	public void ApplyAdjustments(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.Adjustments();
	}

	public void ApplyAutoShapeType(IRibbonControl control)
	{
		PowerPointAddIn1.FormatPainter.Ribbon.AutoShapeType();
	}

	public string MasterShapeMenu(IRibbonControl control)
	{
		return Callbacks.Menu();
	}

	public void InsertTextBox(IRibbonControl control)
	{
		PowerPointAddIn1.MasterShapes.TextBox.Insert();
	}

	public void ToggleMasterShape(ref IRibbonControl control, bool pressed)
	{
		AddRemove.Toggle(control, pressed);
	}

	public void InsertMasterShape(IRibbonControl control)
	{
		AddRemove.Insert(control);
	}

	public bool MasterShapeExists(IRibbonControl control)
	{
		return Callbacks.IsPresent(control);
	}

	public void MasterShapesUpdate(IRibbonControl control)
	{
		PowerPointAddIn1.MasterShapes.Placeholders.Update();
	}

	public void MasterShapesEdit(IRibbonControl control)
	{
		PowerPointAddIn1.MasterShapes.Base.B();
	}

	public string StampsMenu(IRibbonControl control)
	{
		return Stamps.Menu(control);
	}

	public void ToggleStamp(ref IRibbonControl control, bool pressed)
	{
		Stamps.Toggle(control, pressed);
	}

	public bool StampExists(IRibbonControl control)
	{
		return Stamps.IsVisible(control);
	}

	public void StampCustom(IRibbonControl control)
	{
		Stamps.Custom(control);
	}

	public void ToggleSectionTitles(ref IRibbonControl control, bool pressed)
	{
		SectionTitles.Toggle(control, pressed);
	}

	public bool IsShowingSectionTitles(IRibbonControl control)
	{
		return SectionTitles.IsVisible();
	}

	public string StylesMenu(IRibbonControl control)
	{
		return Styles.Menu();
	}

	public void StylesApply(IRibbonControl control)
	{
		Styles.Apply(control);
	}

	public void StylesReset(IRibbonControl control)
	{
		Styles.Reset();
	}

	public void StylesNew(IRibbonControl control)
	{
		Styles.Create();
	}

	public void StylesEdit(IRibbonControl control)
	{
		Styles.Edit();
	}

	public void ConformWidth(IRibbonControl control)
	{
		Conform.Width();
	}

	public void ConformHeight(IRibbonControl control)
	{
		Conform.Height();
	}

	public void ConformBoth(IRibbonControl control)
	{
		Conform.Size();
	}

	public void ConformAdjustments(IRibbonControl control)
	{
		Conform.Adjustments();
	}

	public void ConformPoints(IRibbonControl control)
	{
		Conform.Points();
	}

	public string AlignLastImage(IRibbonControl control)
	{
		return Align.A.A;
	}

	public string AlignLastScreentip(IRibbonControl control)
	{
		return Align.A.B;
	}

	public string AlignLastSupertip(IRibbonControl control)
	{
		return Align.A.C;
	}

	public void AlignLast(IRibbonControl control)
	{
		Align.G();
	}

	public void AlignLeft(IRibbonControl control)
	{
		Align.H();
	}

	public void AlignRight(IRibbonControl control)
	{
		Align.I();
	}

	public void AlignCenter(IRibbonControl control)
	{
		Align.L();
	}

	public void AlignTop(IRibbonControl control)
	{
		Align.J();
	}

	public void AlignBottom(IRibbonControl control)
	{
		Align.K();
	}

	public void AlignMiddle(IRibbonControl control)
	{
		Align.M();
	}

	public void AutoAlign(IRibbonControl control)
	{
		Align.AutoAlign();
	}

	public void AlignOverTable(IRibbonControl control)
	{
		Align.OverTable();
	}

	public Bitmap SwapLastImage(IRibbonControl control)
	{
		return PowerPointAddIn1.Shapes.Swap.A.A;
	}

	public string SwapLastScreentip(IRibbonControl control)
	{
		return PowerPointAddIn1.Shapes.Swap.A.A;
	}

	public string SwapLastSupertip(IRibbonControl control)
	{
		return PowerPointAddIn1.Shapes.Swap.A.B;
	}

	public void SwapLast(IRibbonControl control)
	{
		PowerPointAddIn1.Shapes.Swap.F();
	}

	public void TopLeftAnchor(IRibbonControl control)
	{
		PowerPointAddIn1.Shapes.Swap.G();
	}

	public void TopRightAnchor(IRibbonControl control)
	{
		PowerPointAddIn1.Shapes.Swap.H();
	}

	public void BottomLeftAnchor(IRibbonControl control)
	{
		PowerPointAddIn1.Shapes.Swap.I();
	}

	public void BottomRightAnchor(IRibbonControl control)
	{
		PowerPointAddIn1.Shapes.Swap.J();
	}

	public void CenterAnchor(IRibbonControl control)
	{
		PowerPointAddIn1.Shapes.Swap.K();
	}

	public Bitmap StackLastImage(IRibbonControl control)
	{
		return PowerPointAddIn1.Shapes.Stack.A.A;
	}

	public string StackLastScreentip(IRibbonControl control)
	{
		return PowerPointAddIn1.Shapes.Stack.A.A;
	}

	public string StackLastSupertip(IRibbonControl control)
	{
		return PowerPointAddIn1.Shapes.Stack.A.B;
	}

	public void StackLast(IRibbonControl control)
	{
		PowerPointAddIn1.Shapes.Stack.E();
	}

	public void StackLeft(IRibbonControl control)
	{
		PowerPointAddIn1.Shapes.Stack.F();
	}

	public void StackRight(IRibbonControl control)
	{
		PowerPointAddIn1.Shapes.Stack.G();
	}

	public void StackUp(IRibbonControl control)
	{
		PowerPointAddIn1.Shapes.Stack.H();
	}

	public void StackDown(IRibbonControl control)
	{
		PowerPointAddIn1.Shapes.Stack.I();
	}

	public void DistributeHorizontally(IRibbonControl control)
	{
		Distribute.A();
	}

	public void DistributeVertically(IRibbonControl control)
	{
		Distribute.B();
	}

	public void ConformSize(IRibbonControl control)
	{
		Resize.StandardSize(Conversions.ToInteger(control.Tag));
	}

	public void ConformExcel(IRibbonControl control)
	{
		Resize.A();
	}

	public void ConformWord(IRibbonControl control)
	{
		Resize.B();
	}

	public string ResizeToStandardSizeMenu(IRibbonControl control)
	{
		return Resize.A();
	}

	public void SnapTop(IRibbonControl control)
	{
		SnapToEdge.Top();
	}

	public void SnapBottom(IRibbonControl control)
	{
		SnapToEdge.Bottom();
	}

	public void SnapLeft(IRibbonControl control)
	{
		SnapToEdge.Left();
	}

	public void SnapRight(IRibbonControl control)
	{
		SnapToEdge.Right();
	}

	public void RecordSizePosition(IRibbonControl control)
	{
		Memorize.RecordSizePosition();
	}

	public void RestoreSizePosition(IRibbonControl control)
	{
		Memorize.RestoreSizePosition();
	}

	public string ShowGuideMenu(IRibbonControl control)
	{
		return Guides.ShowGuideMenu();
	}

	public void MultiplyShape(IRibbonControl control)
	{
		Multiply.Shape();
	}

	public void SplitShape(IRibbonControl control)
	{
		Split.Shape();
	}

	public void DuplicateShape(IRibbonControl control)
	{
		PowerPointAddIn1.Shapes.Duplicate.Shape();
	}

	public void ConvertSelectedToPic(IRibbonControl control)
	{
		ConvertToPicture.SelectedObjects();
	}

	public void ConvertEmbeddedToPic(IRibbonControl control)
	{
		ConvertToPicture.AllEmbeddedWorksheets();
	}

	public void ConvertChartsToPic(IRibbonControl control)
	{
		ConvertToPicture.AllCharts();
	}

	public void FixDistortion(IRibbonControl control)
	{
		Images.FixScale();
	}

	public void LineSpacingIncrease(IRibbonControl control)
	{
		LineSpacing.Increase();
	}

	public void LineSpacingDecrease(IRibbonControl control)
	{
		LineSpacing.Decrease();
	}

	public void FootnoteAdd(IRibbonControl control)
	{
		Footnotes.Add();
	}

	public void FootnoteRemove(IRibbonControl control)
	{
		Footnotes.Remove();
	}

	public void FixBullets(IRibbonControl control)
	{
		Bullets.Fix();
	}

	public void SwapText(IRibbonControl control)
	{
		PowerPointAddIn1.TextOps.Swap.SwapText();
	}

	public void MergeText(IRibbonControl control)
	{
		PowerPointAddIn1.TextOps.TextBox.A();
	}

	public void SplitText(IRibbonControl control)
	{
		PowerPointAddIn1.TextOps.TextBox.B();
	}

	public void TextboxMarginsToggle(IRibbonControl control)
	{
		PowerPointAddIn1.TextOps.TextBox.C();
	}

	public void UngroupTable(IRibbonControl control)
	{
		Tables.UngroupTable();
	}

	public void AutofitToggle(IRibbonControl control)
	{
		Autofit.Toggle();
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

	public void ReplaceFonts(IRibbonControl control)
	{
		Fonts.Replace();
	}

	public void BulletArabicPeriod(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletArabicPeriod);
	}

	public void BulletArabicParenRight(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletArabicParenRight);
	}

	public void BulletArabicParenBoth(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletArabicParenBoth);
	}

	public void BulletAlphaUCPeriod(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletAlphaUCPeriod);
	}

	public void BulletAlphaUCParenRight(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletAlphaUCParenRight);
	}

	public void BulletAlphaUCParenBoth(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletAlphaUCParenBoth);
	}

	public void BulletAlphaLCPeriod(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletAlphaLCPeriod);
	}

	public void BulletAlphaLCParenRight(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletAlphaLCParenRight);
	}

	public void BulletAlphaLCParenBoth(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletAlphaLCParenBoth);
	}

	public void BulletRomanUCPeriod(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletRomanUCPeriod);
	}

	public void BulletRomanUCParenRight(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletRomanUCParenRight);
	}

	public void BulletRomanUCParenBoth(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletRomanUCParenBoth);
	}

	public void BulletRomanLCPeriod(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletRomanLCPeriod);
	}

	public void BulletRomanLCParenRight(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletRomanLCParenRight);
	}

	public void BulletRomanLCParenBoth(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletRomanLCParenBoth);
	}

	public void BulletCircleNumDBPlain(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletCircleNumDBPlain);
	}

	public void BulletCircleNumWDBlackPlain(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletCircleNumWDBlackPlain);
	}

	public void BulletCircleNumWDWhitePlain(IRibbonControl control)
	{
		Bullets.ApplyNumberedBullets(MsoNumberedBulletStyle.msoBulletCircleNumWDWhitePlain);
	}

	public void SelectMatchToggle(ref IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.Shapes.SelectMatch.Pane.Toggle(pressed);
	}

	public bool IsSelectMatchOpen(IRibbonControl control)
	{
		return PowerPointAddIn1.Shapes.SelectMatch.Pane.IsVisible();
	}

	public void ArrangeShapesToggle(ref IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.Shapes.Arrange.Pane.Toggle(pressed);
	}

	public bool IsArrangeShapesOpen(IRibbonControl control)
	{
		return PowerPointAddIn1.Shapes.Arrange.Pane.IsVisible();
	}

	public bool IsShapeSelected_Callback(IRibbonControl control)
	{
		return A();
	}

	public bool CanImport_Callback(IRibbonControl control)
	{
		bool result = default(bool);
		try
		{
			int num;
			if (!A())
			{
				if (!B())
				{
					num = 0;
					goto IL_0030;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
			}
			num = (clsRibbon.CallbackSlideView() ? 1 : 0);
			goto IL_0030;
			IL_0030:
			result = (byte)num != 0;
			return result;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public bool IsLinkedItem_Callback(IRibbonControl control)
	{
		return PowerPointAddIn1.Links.Ribbon.IsLinkSelected();
	}

	private bool A()
	{
		bool result = default(bool);
		try
		{
			result = NG.A.Application.ActiveWindow.Selection.ShapeRange != null;
			return result;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private bool B()
	{
		bool result = default(bool);
		try
		{
			result = NG.A.Application.ActiveWindow.Selection.SlideRange != null;
			return result;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public bool IsTocSelected_Callback(IRibbonControl control)
	{
		bool result = false;
		try
		{
			Selection selection = NG.A.Application.ActiveWindow.Selection;
			if (selection.Type == PpSelectionType.ppSelectionSlides)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (selection.SlideRange.Count == 1)
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
					if (PowerPointAddIn1.Slides.Helpers.GetSlideType(selection.SlideRange[1]) == SlideType.TableOfContents)
					{
						result = true;
					}
				}
			}
			selection = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public void MatchWidth(ref IRibbonControl control, bool pressed)
	{
		PB.Settings.ImportMatchDestinationWidth = pressed;
	}

	public void MatchHeight(ref IRibbonControl control, bool pressed)
	{
		PB.Settings.ImportMatchDestinationHeight = pressed;
	}

	public bool GetMatchWidthChecked(IRibbonControl control)
	{
		return PB.Settings.ImportMatchDestinationWidth;
	}

	public bool GetMatchHeightChecked(IRibbonControl control)
	{
		return PB.Settings.ImportMatchDestinationHeight;
	}

	public bool View_Callback(IRibbonControl control)
	{
		clsRibbon.CallbackView(control);
		bool result = default(bool);
		return result;
	}

	public bool PresentationOpen_Callback(IRibbonControl control)
	{
		bool result;
		try
		{
			int num;
			if (IG.A(NG.A.Application.Presentations) != 0)
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
				if (!PowerPointAddIn1.Presentation.Miscellaneous.IsProtectedView(SuppressMessages: true))
				{
					num = (PowerPointAddIn1.Presentation.Miscellaneous.A() ? 1 : 0);
					goto IL_0042;
				}
			}
			num = 0;
			goto IL_0042;
			IL_0042:
			result = (byte)num != 0;
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

	public bool IsGrouped_Callback(IRibbonControl control)
	{
		clsRibbon.CallbackIsGrouped();
		bool result = default(bool);
		return result;
	}

	public bool NotGrouped_Callback(IRibbonControl control)
	{
		clsRibbon.CallbackNotGrouped();
		bool result = default(bool);
		return result;
	}

	public void ViewSlideMaster(IRibbonControl control, bool pressed, ref bool cancelDefault)
	{
		if (!pressed)
		{
			return;
		}
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			cancelDefault = clsCustomize.SlideMaster();
			return;
		}
	}

	public void PictureInsertFromFilePowerPoint_Repurposed(IRibbonControl control, ref bool cancelDefault)
	{
		cancelDefault = PowerPointAddIn1.Library2.Ribbon.PictureOverride();
	}

	public void ChartInsert_Repurposed(IRibbonControl control, ref bool cancelDefault)
	{
		cancelDefault = PowerPointAddIn1.Library2.Ribbon.ChartOverride();
	}

	public void ImportExcel(IRibbonControl control)
	{
		PowerPointAddIn1.Links.ImportExcel.A();
	}

	public void ImportExcelAsGraphic(IRibbonControl control)
	{
		PowerPointAddIn1.Links.ImportExcel.B();
	}

	public void ImportExcelAsImage(IRibbonControl control)
	{
		PowerPointAddIn1.Links.ImportExcel.C();
	}

	public void ImportExcelAsTable(IRibbonControl control)
	{
		PowerPointAddIn1.Links.ImportExcel.D();
	}

	public void ImportExcelAsEmbedded(IRibbonControl control)
	{
		PowerPointAddIn1.Links.ImportExcel.E();
	}

	public void ImportExcelAsText(IRibbonControl control)
	{
		PowerPointAddIn1.Links.ImportExcel.F();
	}

	public void ImportExcelAsChart(IRibbonControl control)
	{
		PowerPointAddIn1.Links.ImportExcel.G();
	}

	public void PrintAreasSheet(IRibbonControl control)
	{
		PagesToPowerPoint.PrintAreasSelectedSheets((Microsoft.Office.Interop.Excel.Application)null, NG.A.Application);
	}

	public void PrintAreasAll(IRibbonControl control)
	{
		PagesToPowerPoint.PrintAreasAllSheets((Microsoft.Office.Interop.Excel.Application)null, NG.A.Application);
	}

	public void Copy_Repurposed(IRibbonControl control, ref bool cancelDefault)
	{
		cancelDefault = false;
	}

	public void LinkedWizard(IRibbonControl control)
	{
		Common.LinkWizard();
	}

	public void EditSource(IRibbonControl control)
	{
		PowerPointAddIn1.Links.Ribbon.EditLinks();
	}

	public void ViewSource(IRibbonControl control)
	{
		PowerPointAddIn1.Links.Ribbon.ViewSource();
	}

	public void UpdateLink(IRibbonControl control)
	{
		PowerPointAddIn1.Links.Ribbon.RefreshLinks();
	}

	public void UpdateAllLinkedShapes(IRibbonControl control)
	{
		PowerPointAddIn1.Links.Shapes.UpdateAllLinks();
	}

	public void RemoveLink(IRibbonControl control)
	{
		PowerPointAddIn1.Links.Ribbon.BreakLinks();
	}

	public void RemoveHyperlinks(IRibbonControl control)
	{
		PowerPointAddIn1.Links.Hyperlinks.RemoveHyperlinks();
	}

	public void HighlightToggle(ref IRibbonControl control, bool pressed)
	{
		Highlight.Toggle(pressed);
	}

	public bool IsHighlighted(IRibbonControl control)
	{
		return Highlight.IsHighlighted;
	}

	public void CreateTemplatedShape(IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.Shapes.Templated.Pane.Toggle(pressed);
	}

	public bool IsCreateTemplatedShapeOpen(IRibbonControl control)
	{
		return PowerPointAddIn1.Shapes.Templated.Pane.IsVisible();
	}

	public void ReapplyLayout(IRibbonControl control)
	{
	}

	public void SlideNumbers(IRibbonControl control)
	{
		Numbers.ShowDialog();
	}

	public void InsertShape(ref IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.Library2.UI.Pane.Toggle(pressed);
	}

	public bool IsContentPaneOpen(IRibbonControl control)
	{
		return PowerPointAddIn1.Library2.UI.Pane.IsVisible();
	}

	public void Ungroup(IRibbonControl control)
	{
		Group.Ungroup();
	}

	public void Regroup(IRibbonControl control)
	{
		Group.Regroup();
	}

	public void TemplateWizard(IRibbonControl control)
	{
		Dialog.Show();
	}

	public void PublishSharedSettings(IRibbonControl control)
	{
		SharedSettings.Publish(AH.A(116727));
	}

	public void Settings(IRibbonControl control)
	{
		if (!Base.ConfigureMacabacus((Microsoft.Office.Interop.Excel.Application)null, NG.A.Application, (Microsoft.Office.Interop.Word.Application)null, KG.A, OB.Ribbon))
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
			KG.A = new clsSettings();
			Create.A();
			PowerPointAddIn1.TurboShapes.Base.ResetColors();
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

	public void DeckCheckToggle(ref IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.DeckCheck.UI.Pane.A(pressed, B: true);
	}

	public void CheckPresentation(IRibbonControl control)
	{
		PowerPointAddIn1.DeckCheck.UI.Pane.B();
	}

	public void CheckSelection(IRibbonControl control)
	{
		PowerPointAddIn1.DeckCheck.UI.Pane.C();
	}

	public bool IsDeckCheckPaneOpen(IRibbonControl control)
	{
		return PowerPointAddIn1.DeckCheck.UI.Pane.A();
	}

	public string ProofingLanguageMenu(IRibbonControl control)
	{
		return PowerPointAddIn1.DeckCheck.Language.LanguagesMenu();
	}

	public void SetLanguage(IRibbonControl control)
	{
		PowerPointAddIn1.DeckCheck.Language.SetProofingLanguage(control);
	}

	public void RecolorToggle(ref IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.Colors.Recolor.Pane.Toggle(pressed);
	}

	public bool IsRecolorOpen(IRibbonControl control)
	{
		return PowerPointAddIn1.Colors.Recolor.Pane.IsVisible();
	}

	public void FixGrayscale(IRibbonControl control)
	{
		Grayscale.A();
	}

	public void EmailDocument(IRibbonControl control)
	{
		Send.ShowDialog();
	}

	public void PdfToFolder(IRibbonControl control)
	{
		Pdf.ToFolder();
	}

	public void PrepareToShare(ref IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.Publishing.Share.Pane.Toggle(pressed);
	}

	public bool IsPrepareToShareOpen(IRibbonControl control)
	{
		return PowerPointAddIn1.Publishing.Share.Pane.IsVisible();
	}

	public void AiwaToggle(ref IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.Aiwa.Pane.Toggle(pressed);
	}

	public bool IsAiwaOpen(IRibbonControl control)
	{
		return PowerPointAddIn1.Aiwa.Pane.IsVisible();
	}

	public bool ShowBetaTools(IRibbonControl control)
	{
		return clsRibbon.ShowBetaTools;
	}

	public void LogoLibraryToggle(ref IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.LogoLibrary.Pane.Toggle(pressed);
	}

	public bool IsLogoLibraryOpen(IRibbonControl control)
	{
		return PowerPointAddIn1.LogoLibrary.Pane.IsVisible();
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

	public string NewPresentationMenu(IRibbonControl control)
	{
		return Templates.BuildNewPresentationMenu();
	}

	public string NewSlideMenu(IRibbonControl control)
	{
		return Templates.BuildNewSlideMenu();
	}

	public string ApplyTemplateMenu(IRibbonControl control)
	{
		return Templates.BuildApplyTemplateMenu();
	}

	public void NewPresentation(IRibbonControl control)
	{
		Create.NewPresentation(control.Tag);
	}

	public void NewSlide(IRibbonControl control)
	{
		Create.NewSlide(control.Tag);
	}

	public void ApplyTemplate(IRibbonControl control)
	{
		Templates.ApplyTemplate(control.Tag);
	}

	public void LibraryShowAll(IRibbonControl control)
	{
		PowerPointAddIn1.Library2.UI.Pane.ShowFromRibbon(blnSlides: true, blnShapes: true, blnImages: true, blnCharts: true, blnText: true, blnDecks: true, blnVideos: true, blnPDFs: true);
	}

	public void LibraryShowSlides(IRibbonControl control)
	{
		PowerPointAddIn1.Library2.UI.Pane.ShowFromRibbon(blnSlides: true, blnShapes: false, blnImages: false, blnCharts: false, blnText: false, blnDecks: false, blnVideos: false);
	}

	public void LibraryShowShapes(IRibbonControl control)
	{
		PowerPointAddIn1.Library2.UI.Pane.ShowFromRibbon(blnSlides: false, blnShapes: true, blnImages: false, blnCharts: false, blnText: false, blnDecks: false, blnVideos: false);
	}

	public void LibraryShowImages(IRibbonControl control)
	{
		PowerPointAddIn1.Library2.UI.Pane.ShowFromRibbon(blnSlides: false, blnShapes: false, blnImages: true, blnCharts: false, blnText: false, blnDecks: false, blnVideos: false);
	}

	public void LibraryShowVideos(IRibbonControl control)
	{
		PowerPointAddIn1.Library2.UI.Pane.ShowFromRibbon(blnSlides: false, blnShapes: false, blnImages: false, blnCharts: false, blnText: false, blnDecks: false, blnVideos: true);
	}

	public void LibraryShowCharts(IRibbonControl control)
	{
		PowerPointAddIn1.Library2.UI.Pane.ShowFromRibbon(blnSlides: false, blnShapes: false, blnImages: false, blnCharts: true, blnText: false, blnDecks: false, blnVideos: false);
	}

	public void LibraryShowText(IRibbonControl control)
	{
		PowerPointAddIn1.Library2.UI.Pane.ShowFromRibbon(blnSlides: false, blnShapes: false, blnImages: false, blnCharts: false, blnText: true, blnDecks: false, blnVideos: false);
	}

	public void LibraryShowArchivedDecks(IRibbonControl control)
	{
		PowerPointAddIn1.Library2.UI.Pane.ShowFromRibbon(blnSlides: false, blnShapes: false, blnImages: false, blnCharts: false, blnText: false, blnDecks: true, blnVideos: false);
	}

	public void ManageLibraryContents(IRibbonControl control)
	{
		Admin.LibraryManager(NG.A.Application, (Microsoft.Office.Interop.Excel.Application)null, (Microsoft.Office.Interop.Word.Application)null);
		InsertSlide.LayoutThumbnails = null;
	}

	public void LibraryVersionControl(IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.Library2.Versioning.Pane.Toggle(pressed);
	}

	public bool IsLibContentPaneOpen(IRibbonControl control)
	{
		return PowerPointAddIn1.Library2.Versioning.Pane.IsVisible();
	}

	public void AirplaneModeToggle(ref IRibbonControl control, bool pressed)
	{
		AirplaneMode.Toggle(pressed);
	}

	public bool IsAirplaneMode(IRibbonControl control)
	{
		return AirplaneMode.IsOn();
	}

	public void AirplanePeek(IRibbonControl control)
	{
		AirplaneMode.Peek();
	}

	public void AirplaneExclude(IRibbonControl control)
	{
		AirplaneMode.Exclude();
	}

	public void AirplaneInclude(IRibbonControl control)
	{
		AirplaneMode.Include();
	}

	public void SaveAll(ref IRibbonControl control)
	{
		Save.All();
	}

	public void SaveUp(IRibbonControl control)
	{
		Save.Up(NG.A.Application.ActivePresentation);
	}

	public void CloseOthers(IRibbonControl control)
	{
		PowerPointAddIn1.Presentation.Miscellaneous.CloseOthers();
	}

	public void Reopen(IRibbonControl control)
	{
		PowerPointAddIn1.Presentation.Miscellaneous.Reopen();
	}

	public void Duplicate(IRibbonControl control)
	{
		PowerPointAddIn1.Presentation.Miscellaneous.Duplicate();
	}

	public void ShowInFolder(IRibbonControl control)
	{
		PowerPointAddIn1.Presentation.Miscellaneous.OpenFolder();
	}

	public void CopyPath(IRibbonControl control)
	{
		PowerPointAddIn1.Presentation.Miscellaneous.CopyPath();
	}

	public void AnalyzeFileSize(IRibbonControl control)
	{
		PowerPointAddIn1.Presentation.Miscellaneous.AnalyzeFileSize();
	}

	public void SendToEnd(IRibbonControl control)
	{
		PowerPointAddIn1.Slides.Miscellaneous.SendToEnd();
	}

	public void RemoveUnusedLayouts(IRibbonControl control)
	{
		PowerPointAddIn1.Slides.Miscellaneous.RemoveUnusedLayouts();
	}

	public void RenameSlide(IRibbonControl control)
	{
		PowerPointAddIn1.Slides.Miscellaneous.Rename();
	}

	public void LockSlides1(IRibbonControl control)
	{
		Protection.LockSlides1();
	}

	public void LockSlides2(IRibbonControl control)
	{
		Protection.LockSlides2();
	}

	public void LockSlides3(IRibbonControl control)
	{
		Protection.LockSlides3();
	}

	public void BackstageShow(object contextObject)
	{
		KG.A = true;
	}

	public void BackstageHide(object contextObject)
	{
		KG.A = false;
	}

	public void ExplorerToggle(ref IRibbonControl control, bool pressed)
	{
		PowerPointAddIn1.Explorer.Pane.Toggle(pressed);
	}

	public bool IsExplorerOpen(IRibbonControl control)
	{
		return PowerPointAddIn1.Explorer.Pane.IsOpen;
	}

	public bool ShowLMSCourseUrl(IRibbonControl control)
	{
		Profile userProfile = Base.UserProfile;
		object value;
		if (userProfile == null)
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
			value = null;
		}
		else
		{
			value = userProfile.LMSCourseUrl;
		}
		return !string.IsNullOrWhiteSpace((string)value);
	}

	public void HelpCenter(IRibbonControl control)
	{
		clsSupport.OnlineDocs(AH.A(167103));
	}

	public void EmailSupport(IRibbonControl control)
	{
		clsSupport.EmailSupport((CallingApp)2);
	}

	public void Feedback(IRibbonControl control)
	{
		Form.Show((OfficeApp)2);
	}

	public string GetSupportDescription(IRibbonControl control)
	{
		return clsSupport.GetSupportDescription();
	}

	public void GoToLMSCourse(IRibbonControl control)
	{
		clsSupport.GoToLMSCourse();
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

	public bool ShowNewerVersionNotice(IRibbonControl control)
	{
		return clsUpdate.ShowNewerVersionNotice(KG.A);
	}

	public void DownloadUpdate(ref IRibbonControl control)
	{
		clsUpdate.DownloadUpdate(KG.A);
	}

	public void DismissUpdate(ref IRibbonControl control)
	{
		clsUpdate.DismissUpdate(KG.A);
	}

	public string UpdateLabel(IRibbonControl control)
	{
		return clsUpdate.NewerVersionLabel();
	}

	public void TagInspector(IRibbonControl control)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		IEnumerator enumerator = default(IEnumerator);
		Microsoft.Office.Interop.PowerPoint.Shape shape = default(Microsoft.Office.Interop.PowerPoint.Shape);
		Tags tags = default(Tags);
		int count = default(int);
		int num5 = default(int);
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 697:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_0036;
						case 4:
							goto IL_0060;
						case 5:
							goto IL_006c;
						case 6:
							goto IL_007e;
						case 7:
							goto IL_00af;
						case 8:
							goto IL_00d0;
						case 9:
							goto IL_00d3;
						case 10:
							goto IL_0106;
						case 11:
							goto IL_0112;
						case 12:
							goto IL_0134;
						case 13:
							goto IL_0167;
						case 14:
							goto IL_017d;
						case 15:
							goto IL_019f;
						case 16:
							goto IL_01d0;
						case 17:
							goto IL_01e8;
						case 18:
							goto IL_0200;
						case 19:
							goto IL_0234;
						case 20:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 21:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_01d3:
					if (enumerator.MoveNext())
					{
						_ = (TextRange)enumerator.Current;
						goto IL_01d0;
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
					goto IL_01e8;
					IL_0007:
					num2 = 2;
					shape = NG.A.Application.ActiveWindow.Selection.ShapeRange[1];
					goto IL_0036;
					IL_0036:
					num2 = 3;
					Interaction.MsgBox(shape.TextFrame2.TextRange.get_Runs(-1, -1).Count);
					goto IL_0060;
					IL_0060:
					num2 = 4;
					tags = shape.Tags;
					goto IL_006c;
					IL_006c:
					num2 = 5;
					count = tags.Count;
					num5 = 1;
					goto IL_00b7;
					IL_00b7:
					if (num5 <= count)
					{
						goto IL_007e;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					goto IL_00d0;
					IL_0106:
					num2 = 10;
					goto IL_0109;
					IL_01e8:
					num2 = 17;
					if (enumerator is IDisposable)
					{
						(enumerator as IDisposable).Dispose();
					}
					goto IL_0200;
					IL_00af:
					num2 = 7;
					num5 = checked(num5 + 1);
					goto IL_00b7;
					IL_00d0:
					tags = null;
					goto IL_00d3;
					IL_00d3:
					num2 = 9;
					enumerator2 = shape.TextFrame2.TextRange.get_Runs(-1, -1).GetEnumerator();
					goto IL_0109;
					IL_0109:
					if (enumerator2.MoveNext())
					{
						_ = (TextRange2)enumerator2.Current;
						goto IL_0106;
					}
					goto IL_0112;
					IL_0112:
					num2 = 11;
					if (enumerator2 is IDisposable)
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
						(enumerator2 as IDisposable).Dispose();
					}
					goto IL_0134;
					IL_0234:
					num2 = 19;
					_ = NG.A.Application.ActivePresentation;
					break;
					IL_007e:
					num2 = 6;
					Interaction.MsgBox(tags.Name(num5) + AH.A(7894) + tags.Value(num5));
					goto IL_00af;
					IL_0134:
					num2 = 12;
					enumerator3 = shape.TextFrame2.TextRange.get_Words(-1, -1).GetEnumerator();
					goto IL_016a;
					IL_016a:
					if (enumerator3.MoveNext())
					{
						_ = (TextRange2)enumerator3.Current;
						goto IL_0167;
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
					goto IL_017d;
					IL_01d0:
					num2 = 16;
					goto IL_01d3;
					IL_017d:
					num2 = 14;
					if (enumerator3 is IDisposable)
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
						(enumerator3 as IDisposable).Dispose();
					}
					goto IL_019f;
					IL_0167:
					num2 = 13;
					goto IL_016a;
					IL_0200:
					num2 = 18;
					_ = NG.A.Application.ActiveWindow.Selection.SlideRange[1];
					goto IL_0234;
					IL_019f:
					num2 = 15;
					enumerator = shape.TextFrame.TextRange.Words().GetEnumerator();
					goto IL_01d3;
					end_IL_0000_2:
					break;
				}
				num2 = 20;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 697;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
	}

	private object A()
	{
		throw new NotImplementedException();
	}
}
