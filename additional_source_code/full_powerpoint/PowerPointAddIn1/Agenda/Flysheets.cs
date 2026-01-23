using System;
using System.Collections;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Explorer;
using PowerPointAddIn1.MasterShapes;
using PowerPointAddIn1.Slides;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1.Agenda;

public sealed class Flysheets
{
	public static Slide InsertFlysheet()
	{
		Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
		Slide slide = null;
		string strTopic = "";
		int intLevel = 0;
		bool blnCancel = false;
		CustomLayout customLayout = null;
		try
		{
			presentation = application.ActivePresentation;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (presentation != null)
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
			FlySheetStyle presentationFlysheetStyle = Behavior.GetPresentationFlysheetStyle(presentation);
			SlideType type;
			if (presentationFlysheetStyle == FlySheetStyle.Topic)
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
				type = SlideType.Flysheet;
			}
			else
			{
				type = SlideType.Agenda;
			}
			customLayout = Helpers.GetLayout(presentation, type);
			if (customLayout != null)
			{
				PromptForSectionTitle(ref strTopic, ref intLevel, ref blnCancel);
				if (!blnCancel)
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
					application.StartNewUndoEntry();
					slide = presentation.Slides.AddSlide(checked(Helpers.GetSlideIndex() + 1), customLayout);
					slide.Select();
					Create.SkipPresentationNewSlideEvent(slide);
					presentation.SectionProperties.AddBeforeSlide(slide.SlideIndex, (intLevel == 1) ? strTopic : (Constants.SUBSECTION_PREFIX + strTopic));
					if (presentationFlysheetStyle == FlySheetStyle.Agenda)
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
						UseExistingAgendaSlideTitle(presentation, slide);
						Update.RefreshAgendaStyle(presentation);
					}
					else
					{
						Update.RefreshTopicStyle(presentation);
					}
				}
				SectionTitles.Update(presentation, SuppressMessages: true);
				customLayout = null;
			}
			else
			{
				Forms.WarningMessage(AH.A(4422));
			}
			presentation = null;
		}
		application = null;
		return slide;
	}

	public static void UseExistingAgendaSlideTitle(Microsoft.Office.Interop.PowerPoint.Presentation pres, Slide sldNew)
	{
		IEnumerator enumerator = pres.Slides.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Slide slide = (Slide)enumerator.Current;
				if (slide.SlideIndex == sldNew.SlideIndex)
				{
					continue;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (Helpers.GetSlideType(slide) != SlideType.Agenda)
				{
					continue;
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
				if (slide.Shapes.HasTitle != MsoTriState.msoTrue)
				{
					continue;
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
				Microsoft.Office.Interop.PowerPoint.Shape title = slide.Shapes.Title;
				if (title.HasTextFrame == MsoTriState.msoTrue && title.TextFrame2.HasText == MsoTriState.msoTrue)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							try
							{
								sldNew.Shapes.Title.TextFrame2.TextRange.Text = title.TextFrame2.TextRange.Text;
								return;
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
								return;
							}
						}
					}
				}
				title = null;
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	public static void PromptForSectionTitle(ref string strTopic, ref int intLevel, ref bool blnCancel)
	{
		wpfSectionAdd wpfSectionAdd2 = new wpfSectionAdd();
		wpfSectionAdd2.ShowDialog();
		if (wpfSectionAdd2.DialogResult.HasValue)
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
			if (wpfSectionAdd2.DialogResult.Value)
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
				strTopic = wpfSectionAdd2.txtTopic.Text;
				if (wpfSectionAdd2.radSubsection.IsChecked == true)
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
					strTopic = Strings.Mid(strTopic, 2);
					intLevel = 2;
				}
				else
				{
					intLevel = 1;
				}
				goto IL_009e;
			}
		}
		blnCancel = true;
		goto IL_009e;
		IL_009e:
		wpfSectionAdd2 = null;
	}

	public static void FlysheetStyleTopic()
	{
		Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
		application.StartNewUndoEntry();
		_ = null;
		A(activePresentation, SlideType.Agenda);
		Behavior.SetPresentationFlysheetStyle(activePresentation, FlySheetStyle.Topic);
		if (activePresentation.SectionProperties.Count > 0)
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
				Update.RefreshTopicStyle(activePresentation);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(AH.A(4688) + ex2.Message);
				ProjectData.ClearProjectError();
			}
		}
		activePresentation = null;
		A();
	}

	public static void FlysheetStyleAgenda()
	{
		Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
		application.StartNewUndoEntry();
		_ = null;
		A(activePresentation, SlideType.Flysheet);
		Behavior.SetPresentationFlysheetStyle(activePresentation, FlySheetStyle.Agenda);
		if (activePresentation.SectionProperties.Count > 0)
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
			try
			{
				Update.RefreshAgendaStyle(activePresentation);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(AH.A(4688) + ex2.Message);
				ProjectData.ClearProjectError();
			}
		}
		activePresentation = null;
		A();
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, SlideType B)
	{
		int num = 1;
		checked
		{
			if (B == SlideType.Agenda)
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
				Slide slide = TableOfContents.Slide(A);
				if (slide != null)
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
					num = slide.SlideIndex + 1;
				}
			}
			Events.Disable(A.Application);
			int count = A.Slides.Count;
			int num2 = num;
			Slide slide2;
			for (int i = count; i >= num2; i += -1)
			{
				slide2 = A.Slides[i];
				if (Helpers.GetSlideType(slide2) != B)
				{
					continue;
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
				slide2.Delete();
			}
			slide2 = null;
		}
	}

	public static bool ShowTocMenu()
	{
		bool result;
		try
		{
			result = Behavior.GetPresentationFlysheetStyle(NG.A.Application.ActivePresentation) == FlySheetStyle.Topic;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = true;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static void A()
	{
		KG.A.InvalidateControl(AH.A(4795));
		KG.A.InvalidateControl(AH.A(4822));
	}

	public static void FlysheetLevel(int intLevel)
	{
		if (!Licensing.AllowAgendaOperation())
		{
			return;
		}
		Application application = NG.A.Application;
		try
		{
			Slide slide = application.ActiveWindow.Selection.SlideRange[1];
			SlideType slideType = Helpers.GetSlideType(slide);
			if (slideType == SlideType.Flysheet)
			{
				goto IL_006f;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (slideType == SlideType.Agenda)
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
				goto IL_006f;
			}
			B();
			goto end_IL_0019;
			IL_006f:
			Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
			SectionProperties sectionProperties = activePresentation.SectionProperties;
			int sectionIndex = slide.sectionIndex;
			string text = sectionProperties.Name(sectionIndex);
			if (intLevel == 2)
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
				if (!Update.A(text))
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
					application.StartNewUndoEntry();
					sectionProperties.Rename(sectionIndex, Constants.SUBSECTION_PREFIX + text);
					Update.A(A: true, activePresentation);
					goto IL_011a;
				}
			}
			if (intLevel == 1)
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
				if (Update.A(text))
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
					application.StartNewUndoEntry();
					sectionProperties.Rename(sectionIndex, Strings.Mid(text, 2));
					Update.A(A: true, activePresentation);
				}
			}
			goto IL_011a;
			IL_011a:
			sectionProperties = null;
			activePresentation = null;
			end_IL_0019:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			B();
			ProjectData.ClearProjectError();
		}
		application = null;
	}

	private static void B()
	{
		Forms.WarningMessage(AH.A(4859));
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape GetFlysheetTopicShape(Slide sld)
	{
		Microsoft.Office.Interop.PowerPoint.Shape result = null;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = sld.Shapes.Placeholders.GetEnumerator();
			while (true)
			{
				if (enumerator.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
					Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
					if (shape2.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody)
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
						if (shape2.HasTextFrame == MsoTriState.msoTrue)
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
							if (shape2.TextFrame2.TextRange.ParagraphFormat.Bullet.Type == MsoBulletType.msoBulletNumbered)
							{
								result = shape;
								break;
							}
							_ = null;
						}
					}
					shape2 = null;
					continue;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_0097;
					}
					continue;
					end_IL_0097:
					break;
				}
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		return result;
	}
}
