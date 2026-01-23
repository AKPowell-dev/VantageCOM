using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.Config.Settings;
using MacabacusMacros.Libraries;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Agenda;
using PowerPointAddIn1.Library2;
using PowerPointAddIn1.MasterShapes;
using PowerPointAddIn1.Presentation;
using PowerPointAddIn1.Shapes;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Template;

public sealed class Create
{
	[CompilerGenerated]
	private static bool m_A;

	private static Microsoft.Office.Interop.PowerPoint.Presentation m_A;

	private static Dictionary<string, List<Bitmap>> m_A;

	[CompilerGenerated]
	private static List<int> m_A;

	[CompilerGenerated]
	private static Dictionary<string, bool> m_A;

	[CompilerGenerated]
	private static string m_A;

	private static string B = string.Empty;

	[CompilerGenerated]
	private static List<int> B;

	private static bool HasNewPresentationBeenCreated
	{
		[CompilerGenerated]
		get
		{
			return Create.m_A;
		}
		[CompilerGenerated]
		set
		{
			Create.m_A = value;
		}
	} = false;

	private static List<int> InsertedSlideIds
	{
		[CompilerGenerated]
		get
		{
			return Create.m_A;
		}
		[CompilerGenerated]
		set
		{
			Create.m_A = value;
		}
	}

	private static Dictionary<string, bool> TemplateHasSlides
	{
		[CompilerGenerated]
		get
		{
			return Create.m_A;
		}
		[CompilerGenerated]
		set
		{
			Create.m_A = value;
		}
	}

	internal static string DefaultTemplateId
	{
		[CompilerGenerated]
		get
		{
			return Create.m_A;
		}
		[CompilerGenerated]
		set
		{
			Create.m_A = value;
		}
	} = "";

	private static string A
	{
		get
		{
			if (Operators.CompareString(Create.B, string.Empty, TextCompare: false) == 0)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Create.B = Templates.A(DefaultTemplateId);
			}
			return Create.B;
		}
		set
		{
			Create.B = value;
		}
	}

	private static List<int> ProcessedNewSlideIds
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public static Microsoft.Office.Interop.PowerPoint.Presentation NewBlankPresentation(Microsoft.Office.Interop.PowerPoint.Application ppApp)
	{
		Microsoft.Office.Interop.PowerPoint.Application target = ppApp;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105519)).RemoveEventHandler(target, new EApplication_NewPresentationEventHandler(Application_NewPresentation));
		Microsoft.Office.Interop.PowerPoint.Presentation result = ppApp.Presentations.Add();
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105519)).AddEventHandler(target, new EApplication_NewPresentationEventHandler(Application_NewPresentation));
		target = null;
		return result;
	}

	public static void Application_NewPresentation(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
	{
		if (KG.A)
		{
			if (Pres.Windows.Count <= 0)
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
				clsRibbon.InvalidateOpenPresentationRequiredControls();
				return;
			}
		}
		if (Pres.Windows.Count <= 0)
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			if (!HasNewPresentationBeenCreated)
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
				if (A())
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							Create.m_A = Pres;
							if (A("") == null)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										break;
									default:
										Create.m_A = null;
										return;
									}
								}
							}
							return;
						}
					}
				}
			}
			clsRibbon.InvalidateOpenPresentationRequiredControls();
			return;
		}
	}

	public static void Application_AfterNewPresentation(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
	{
		if (A())
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
			HasNewPresentationBeenCreated = true;
			try
			{
				if (Pres == Create.m_A)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						Pres.Saved = MsoTriState.msoTrue;
						Pres.Close();
						Create.m_A = null;
						break;
					}
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105550)).RemoveEventHandler(NG.A.Application, new EApplication_AfterNewPresentationEventHandler(Application_AfterNewPresentation));
	}

	public static void FileNew_Repurposed(ref bool cancelDefault)
	{
		if (A())
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					cancelDefault = A("") != null;
					return;
				}
			}
		}
		cancelDefault = false;
	}

	private static bool A()
	{
		return DefaultTemplateId.Length > 0;
	}

	public static void NewPresentation(string strTemplatePath)
	{
		if (!Access.AllowPowerPointOperation((PlanType)5, (Restriction)0, false))
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (A(strTemplatePath) == null)
			{
				return;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)11, AH.A(135018));
				return;
			}
		}
	}

	public static void NewSlide(string strTemplatePath)
	{
		if (!Access.AllowPowerPointOperation((PlanType)5, (Restriction)0, false))
		{
			return;
		}
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
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
				Microsoft.Office.Interop.PowerPoint.Presentation presentation = A(strTemplatePath, B: true);
				if (presentation == null)
				{
					return;
				}
				Microsoft.Office.Interop.PowerPoint.Presentation presentation2 = presentation;
				try
				{
					SectionProperties sectionProperties = presentation2.SectionProperties;
					for (int i = sectionProperties.Count; i >= 2; i += -1)
					{
						sectionProperties.Delete(i, deleteSlides: true);
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						sectionProperties = null;
						break;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				try
				{
					for (int j = presentation2.Slides.Count; j >= 1; j += -1)
					{
						presentation2.Slides[j].Delete();
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				try
				{
					enumerator = presentation2.SlideMaster.CustomLayouts.GetEnumerator();
					while (true)
					{
						if (enumerator.MoveNext())
						{
							CustomLayout customLayout = (CustomLayout)enumerator.Current;
							if (PowerPointAddIn1.Slides.Helpers.GetLayoutType(customLayout) != SlideType.Content)
							{
								continue;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								presentation2.Slides.AddSlide(1, customLayout);
								break;
							}
							break;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_011d;
							}
							continue;
							end_IL_011d:
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
							switch (6)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				presentation2 = null;
				clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)11, AH.A(135079));
				return;
			}
		}
	}

	private static Microsoft.Office.Interop.PowerPoint.Presentation A(string A, bool B = false)
	{
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
		Microsoft.Office.Interop.PowerPoint.Presentation result;
		bool c = default(bool);
		if (A.Length == 0)
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
			try
			{
				A = Create.A;
				if (A.Length == 0)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						if (Operators.CompareString(DefaultTemplateId, Sample.SampleTemplateId(), TextCompare: false) == 0)
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
							try
							{
								XmlDocument xml = Manage.GetXml(false);
								xml.DocumentElement.SelectSingleNode(Constants.XML_DEFAULT_TEMPLATE_ID).InnerText = "";
								Manage.Save(xml, true);
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
						}
						DefaultTemplateId = "";
						result = null;
						break;
					}
					goto IL_03ef;
				}
				c = true;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Forms.ErrorMessage(AH.A(135142));
				result = null;
				ProjectData.ClearProjectError();
				goto IL_03ef;
			}
		}
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		bool flag = default(bool);
		try
		{
			flag = AirplaneMode.IsOn();
			if (flag)
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
				AirplaneMode.RemoveEvents(application);
			}
			Microsoft.Office.Interop.PowerPoint.Application target = application;
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(58943)).RemoveEventHandler(target, new EApplication_PresentationNewSlideEventHandler(Application_PresentationNewSlide));
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105519)).RemoveEventHandler(target, new EApplication_NewPresentationEventHandler(Application_NewPresentation));
			if (!HasNewPresentationBeenCreated)
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
				new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105550)).RemoveEventHandler(target, new EApplication_AfterNewPresentationEventHandler(Application_AfterNewPresentation));
			}
			try
			{
				presentation = Create.A(application, A);
			}
			finally
			{
				new ComAwareEventInfo(typeof(EApplication_Event), AH.A(58943)).AddEventHandler(target, new EApplication_PresentationNewSlideEventHandler(Application_PresentationNewSlide));
				new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105519)).AddEventHandler(target, new EApplication_NewPresentationEventHandler(Application_NewPresentation));
				if (!HasNewPresentationBeenCreated)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105550)).AddEventHandler(target, new EApplication_AfterNewPresentationEventHandler(Application_AfterNewPresentation));
						break;
					}
				}
			}
			target = null;
			Templates.A(presentation, A, c);
			clsRibbon.InvalidateOpenPresentationRequiredControls();
			if (!B)
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
				if (presentation.Slides.Count == 1)
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
					if (!Create.A(application, A))
					{
						presentation.Slides[1].Delete();
					}
				}
			}
			try
			{
				string innerText = KG.A.SettingsXml.GetElementsByTagName(Constants.XML_NEW_PRESENTATION_DEFAULT_STAMP).Item(0).InnerText;
				if (innerText.Length == 0)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						Stamps.HideLegacyStamps();
						break;
					}
				}
				else
				{
					Stamps.AddToNewPresentation(innerText);
				}
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			if (flag)
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
				try
				{
					AirplaneMode.HidePresentationImages(presentation);
				}
				catch (Exception ex7)
				{
					ProjectData.SetProjectError(ex7);
					Exception ex8 = ex7;
					ProjectData.ClearProjectError();
				}
			}
			presentation.Saved = MsoTriState.msoTrue;
		}
		catch (OutOfMemoryException ex9)
		{
			ProjectData.SetProjectError(ex9);
			OutOfMemoryException ex10 = ex9;
			Forms.ErrorMessage(ex10.Message);
			ProjectData.ClearProjectError();
		}
		catch (Exception ex11)
		{
			ProjectData.SetProjectError(ex11);
			Exception ex12 = ex11;
			if (!clsFile.IsPathUrl(A) && File.Exists(A))
			{
				if (ex12.Message.Contains(AH.A(135624)))
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
					Forms.ErrorMessage(AH.A(135691));
				}
				else
				{
					Forms.ErrorMessage(ex12.Message);
					clsReporting.LogException(ex12);
				}
			}
			else
			{
				Forms.ErrorMessage(AH.A(136123));
			}
			ProjectData.ClearProjectError();
		}
		finally
		{
			if (flag)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					AirplaneMode.AddEvents(application);
					break;
				}
			}
		}
		application = null;
		result = presentation;
		goto IL_03ef;
		IL_03ef:
		return result;
	}

	internal static void A()
	{
		Create.A = string.Empty;
	}

	private static bool A(Microsoft.Office.Interop.PowerPoint.Application A, string B)
	{
		bool value = false;
		if (TemplateHasSlides == null)
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
			TemplateHasSlides = new Dictionary<string, bool>();
		}
		if (TemplateHasSlides.TryGetValue(B, out value))
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					return value;
				}
			}
		}
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
		try
		{
			presentation = PowerPointAddIn1.Presentation.Helpers.OpenQuietly(A, B);
			value = presentation.Slides.Count > 0;
			TemplateHasSlides.Add(B, value);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		finally
		{
			if (presentation != null)
			{
				try
				{
					PowerPointAddIn1.Presentation.Helpers.CloseQuietly(presentation);
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				presentation = null;
			}
		}
		return value;
	}

	private static Microsoft.Office.Interop.PowerPoint.Presentation A(Microsoft.Office.Interop.PowerPoint.Application A, string B)
	{
		return A.Presentations.Open(B, MsoTriState.msoFalse, MsoTriState.msoTrue);
	}

	public static void ApplyHeadersFooters(Microsoft.Office.Interop.PowerPoint.Presentation pres, Microsoft.Office.Interop.PowerPoint.Presentation template)
	{
		try
		{
			pres.SlideMaster.HeadersFooters.SlideNumber.Visible = template.SlideMaster.HeadersFooters.SlideNumber.Visible;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		try
		{
			pres.SlideMaster.HeadersFooters.Footer.Visible = template.SlideMaster.HeadersFooters.Footer.Visible;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
		try
		{
			pres.SlideMaster.HeadersFooters.DateAndTime.Visible = template.SlideMaster.HeadersFooters.DateAndTime.Visible;
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			clsReporting.LogException(ex6);
			ProjectData.ClearProjectError();
		}
	}

	public static Slide InsertTitleSlide()
	{
		return A(SlideType.Title, AH.A(136395));
	}

	public static Slide InsertLegalSlide()
	{
		return A(SlideType.Legal, AH.A(136406));
	}

	public static Slide InsertContactSlide()
	{
		return A(SlideType.Contact, AH.A(136433));
	}

	public static Slide InsertBlankSlide()
	{
		return A(SlideType.Blank, AH.A(136472));
	}

	public static Slide InsertFrontCoverSlide()
	{
		return A(SlideType.CoverFront, AH.A(136511));
	}

	public static Slide InsertBackCoverSlide()
	{
		return A(SlideType.CoverBack, AH.A(136534));
	}

	private static Slide A(SlideType A, string B)
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
		Slide slide = null;
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
		checked
		{
			if (presentation != null)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				customLayout = PowerPointAddIn1.Slides.Helpers.GetLayout(presentation, A);
				if (customLayout != null)
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
					int count = presentation.Slides.Count;
					int index = default(int);
					switch (A)
					{
					case SlideType.Title:
						if (count > 0)
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
							if (PowerPointAddIn1.Slides.Helpers.GetSlideType(presentation.Slides[1]) == SlideType.CoverFront)
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
								index = 2;
								break;
							}
						}
						index = 1;
						break;
					case SlideType.Contact:
						if (count > 0)
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
							if (PowerPointAddIn1.Slides.Helpers.GetSlideType(presentation.Slides[count]) == SlideType.CoverBack)
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
								index = count;
								break;
							}
						}
						index = count + 1;
						break;
					case SlideType.Legal:
					{
						index = count + 1;
						int num = Math.Max(count - 2, 1);
						for (int i = count; i >= num; i += -1)
						{
							if (PowerPointAddIn1.Slides.Helpers.GetSlideType(presentation.Slides[i]) == SlideType.Contact || PowerPointAddIn1.Slides.Helpers.GetSlideType(presentation.Slides[i]) == SlideType.CoverBack)
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
							index = i + 1;
							break;
						}
						break;
					}
					case SlideType.Blank:
						index = PowerPointAddIn1.Slides.Helpers.GetSlideIndex() + 1;
						break;
					case SlideType.CoverFront:
						index = 1;
						break;
					case SlideType.CoverBack:
						index = count + 1;
						break;
					}
					application.StartNewUndoEntry();
					slide = presentation.Slides.AddSlide(index, customLayout);
					slide.Select();
					customLayout = null;
				}
				else
				{
					Forms.WarningMessage(AH.A(136555) + B + AH.A(136562));
				}
				presentation = null;
			}
			application = null;
			return slide;
		}
	}

	public static Slide InsertTocSlide()
	{
		return TableOfContents.A(A: false);
	}

	public static void SkipPresentationNewSlideEvent(Slide sld)
	{
		if (InsertedSlideIds == null)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			InsertedSlideIds = new List<int>();
		}
		InsertedSlideIds.Add(sld.SlideID);
	}

	public static void Application_PresentationNewSlide(Slide Sld)
	{
		int slideID;
		try
		{
			slideID = Sld.SlideID;
		}
		catch (object obj) when (((Func<bool>)delegate
		{
			// Could not convert BlockContainer to single expression
			Exception obj2 = obj as Exception;
			System.Runtime.CompilerServices.Unsafe.SkipInit(out int result);
			if (obj2 == null)
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
				result = 0;
			}
			else
			{
				ProjectData.SetProjectError(obj2);
				result = ((obj2.HResult == -2147188720) ? 1 : 0);
			}
			return (byte)result != 0;
		}).Invoke())
		{
			ProjectData.ClearProjectError();
			return;
		}
		if (ProcessedNewSlideIds == null)
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
			ProcessedNewSlideIds = new List<int>();
		}
		List<int> processedNewSlideIds = ProcessedNewSlideIds;
		if (!processedNewSlideIds.Contains(slideID))
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					processedNewSlideIds.Add(slideID);
					processedNewSlideIds = null;
					if (((Microsoft.Office.Interop.PowerPoint.Presentation)Sld.Parent).Windows.Count != 0)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								if (InsertedSlideIds != null)
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
									if (InsertedSlideIds.Contains(Sld.SlideID))
									{
										while (true)
										{
											switch (3)
											{
											case 0:
												break;
											default:
												InsertedSlideIds.Remove(Sld.SlideID);
												return;
											}
										}
									}
								}
								A(Sld);
								return;
							}
						}
					}
					return;
				}
			}
		}
		processedNewSlideIds.Remove(slideID);
	}

	private static void A(Slide A)
	{
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = NG.A.Application.ActivePresentation;
		SlideType slideType = PowerPointAddIn1.Slides.Helpers.GetSlideType(A);
		bool blnCancel = false;
		string strTopic = "";
		int intLevel = 0;
		switch (slideType)
		{
		case SlideType.Flysheet:
			if (PowerPointAddIn1.Agenda.Behavior.GetPresentationFlysheetStyle(activePresentation) == FlySheetStyle.Topic)
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
				Microsoft.Office.Interop.PowerPoint.Shape flysheetTopicShape = Flysheets.GetFlysheetTopicShape(A);
				if (flysheetTopicShape == null)
				{
					break;
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
				if (flysheetTopicShape.TextFrame2.HasText == MsoTriState.msoFalse)
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
					Flysheets.PromptForSectionTitle(ref strTopic, ref intLevel, ref blnCancel);
					if (!blnCancel)
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
						TextRange2 textRange = flysheetTopicShape.TextFrame2.TextRange;
						textRange.Text = strTopic;
						textRange.ParagraphFormat.IndentLevel = intLevel;
						_ = null;
					}
				}
				else
				{
					TextRange2 textRange2 = flysheetTopicShape.TextFrame2.TextRange;
					strTopic = textRange2.Text;
					intLevel = textRange2.ParagraphFormat.IndentLevel;
					_ = null;
				}
				if (!blnCancel)
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
					SectionProperties sectionProperties = activePresentation.SectionProperties;
					int slideIndex = A.SlideIndex;
					string sectionName;
					if (intLevel != 1)
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
						sectionName = Constants.SUBSECTION_PREFIX + strTopic;
					}
					else
					{
						sectionName = strTopic;
					}
					sectionProperties.AddBeforeSlide(slideIndex, sectionName);
					try
					{
						Update.RefreshTopicStyle(activePresentation);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				else
				{
					A.Delete();
				}
				flysheetTopicShape = null;
			}
			else
			{
				Forms.WarningMessage(AH.A(136818));
				A.Delete();
			}
			break;
		case SlideType.Title:
		case SlideType.Legal:
		case SlideType.Contact:
		case SlideType.CoverFront:
		case SlideType.CoverBack:
		{
			string text = "";
			IEnumerator enumerator2 = default(IEnumerator);
			try
			{
				enumerator2 = activePresentation.Slides.GetEnumerator();
				while (true)
				{
					if (enumerator2.MoveNext())
					{
						Slide slide = (Slide)enumerator2.Current;
						if (slide == A || PowerPointAddIn1.Slides.Helpers.GetSlideType(slide) != slideType)
						{
							continue;
						}
						switch (slideType)
						{
						default:
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								break;
							}
							if (slideType != SlideType.CoverBack)
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
								text = AH.A(136534);
							}
							break;
						case SlideType.Title:
							text = AH.A(136395);
							break;
						case SlideType.Legal:
							text = AH.A(136406);
							break;
						case SlideType.Contact:
							text = AH.A(136433);
							break;
						case SlideType.CoverFront:
							text = AH.A(136511);
							break;
						case SlideType.TableOfContents:
						case SlideType.Flysheet:
							break;
						}
						if (MessageBox.Show(AH.A(137054) + text + AH.A(137127), AH.A(5874), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								slide.Delete();
								break;
							}
						}
						else
						{
							A.Delete();
						}
						break;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_02d3;
						}
						continue;
						end_IL_02d3:
						break;
					}
					break;
				}
			}
			finally
			{
				if (enumerator2 is IDisposable)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						(enumerator2 as IDisposable).Dispose();
						break;
					}
				}
			}
			break;
		}
		default:
		{
			PowerPointAddIn1.Slides.Helpers.DesignateSlideAsType(A, SlideType.Content);
			if (SectionTitles.IsVisible(activePresentation))
			{
				SectionTitles.Update(activePresentation, SuppressMessages: false, A);
				break;
			}
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.CustomLayout.Shapes.GetEnumerator();
				while (true)
				{
					if (enumerator.MoveNext())
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
						if (!SectionTitles.IsLegacySectionTitlePlaceholder(shape))
						{
							continue;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							shape.Visible = MsoTriState.msoFalse;
							break;
						}
						break;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_0373;
						}
						continue;
						end_IL_0373:
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
						switch (3)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			break;
		}
		case SlideType.TableOfContents:
		case SlideType.Agenda:
		case SlideType.Blank:
			break;
		}
		activePresentation = null;
	}
}
