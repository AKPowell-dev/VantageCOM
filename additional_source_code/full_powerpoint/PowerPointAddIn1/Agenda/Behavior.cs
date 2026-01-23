using System;
using System.Collections;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Agenda;

public sealed class Behavior
{
	private static readonly string m_A = AH.A(5530);

	private static readonly string m_B = AH.A(5567);

	private static readonly string m_C = AH.A(5604);

	private static readonly string m_D = AH.A(5633);

	private static readonly string E = AH.A(5662);

	private static readonly string F = AH.A(5699);

	public static void SetPresentationFlysheetStyle(Microsoft.Office.Interop.PowerPoint.Presentation pres, FlySheetStyle style)
	{
		Tags tags = pres.Tags;
		string f = F;
		int num = (int)style;
		tags.Add(f, num.ToString());
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		bool flag = false;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = pres.Designs.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			do
			{
				if (enumerator.MoveNext())
				{
					Design design = (Design)enumerator.Current;
					{
						enumerator2 = design.SlideMaster.CustomLayouts.GetEnumerator();
						try
						{
							while (true)
							{
								if (enumerator2.MoveNext())
								{
									CustomLayout customLayout = (CustomLayout)enumerator2.Current;
									shape = Helpers.GetSlideTypeShape(customLayout.Shapes);
									if (shape == null)
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
										break;
									}
									if (1 == 0)
									{
										/*OpCode not supported: LdMemberToken*/;
									}
									try
									{
										switch ((SlideType)Conversions.ToInteger(shape.Tags[Helpers.TAG_SLIDE_TYPE].ToString()))
										{
										default:
											while (true)
											{
												switch (1)
												{
												case 0:
													break;
												default:
													goto end_IL_00c7;
												}
												continue;
												end_IL_00c7:
												break;
											}
											goto IL_0124;
										case SlideType.TableOfContents:
											if (style != FlySheetStyle.Agenda)
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
												shape.Delete();
												Helpers.DesignateLayoutAsType(customLayout, SlideType.Agenda);
												flag = true;
												break;
											}
											break;
										case SlideType.Agenda:
											if (style != FlySheetStyle.Topic)
											{
												break;
											}
											while (true)
											{
												switch (7)
												{
												case 0:
													continue;
												}
												shape.Delete();
												Helpers.DesignateLayoutAsType(customLayout, SlideType.TableOfContents);
												flag = true;
												break;
											}
											break;
										}
									}
									catch (Exception ex)
									{
										ProjectData.SetProjectError(ex);
										Exception ex2 = ex;
										ProjectData.ClearProjectError();
										goto IL_0124;
									}
									break;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_0132;
									}
									continue;
									end_IL_0132:
									break;
								}
								break;
								IL_0124:
								shape = null;
							}
						}
						finally
						{
							IDisposable disposable = enumerator2 as IDisposable;
							if (disposable != null)
							{
								disposable.Dispose();
							}
						}
					}
					continue;
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_0166;
					}
					continue;
					end_IL_0166:
					break;
				}
				break;
			}
			while (!flag);
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		IRibbonUI a = KG.A;
		a.InvalidateControl(AH.A(5284));
		a.InvalidateControl(AH.A(5311));
		a.InvalidateControl(AH.A(5340));
		a.InvalidateControl(AH.A(5373));
		a.InvalidateControl(AH.A(5398));
		_ = null;
	}

	public static FlySheetStyle GetPresentationFlysheetStyle(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		FlySheetStyle flySheetStyle = default(FlySheetStyle);
		try
		{
			string text = pres.Tags[F];
			if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
			{
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
					flySheetStyle = (FlySheetStyle)Conversions.ToInteger(KG.A.SettingsXml.GetElementsByTagName(AH.A(5423)).Item(0).InnerText);
					_ = null;
					SetPresentationFlysheetStyle(pres, flySheetStyle);
					break;
				}
			}
			else
			{
				flySheetStyle = (FlySheetStyle)Conversions.ToInteger(text);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return flySheetStyle;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, bool B)
	{
		A.Tags.Add(Behavior.m_B, (B ? 1 : 0).ToString());
	}

	public static bool GetPresentationAutoCollapse(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		bool flag = default(bool);
		try
		{
			string text = pres.Tags[Behavior.m_B];
			if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
			{
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
					flag = Conversions.ToBoolean(KG.A.SettingsXml.GetElementsByTagName(AH.A(5450)).Item(0).InnerText);
					_ = null;
					A(pres, flag);
					break;
				}
			}
			else
			{
				flag = Conversions.ToBoolean(text);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return flag;
	}

	private static void B(Microsoft.Office.Interop.PowerPoint.Presentation A, bool B)
	{
		A.Tags.Add(Behavior.m_C, (B ? 1 : 0).ToString());
	}

	public static bool GetPresentationSkipDoubles(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		bool flag = default(bool);
		try
		{
			string text = pres.Tags[Behavior.m_C];
			if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
			{
				flag = Conversions.ToBoolean(KG.A.SettingsXml.GetElementsByTagName(AH.A(3467)).Item(0).InnerText);
				_ = null;
				B(pres, flag);
			}
			else
			{
				flag = Conversions.ToBoolean(text);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return flag;
	}

	private static void C(Microsoft.Office.Interop.PowerPoint.Presentation A, bool B)
	{
		A.Tags.Add(Behavior.m_D, (B ? 1 : 0).ToString());
	}

	public static bool GetPresentationOmitDoubles(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		bool flag = default(bool);
		try
		{
			string text = pres.Tags[Behavior.m_D];
			if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
			{
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
					flag = Conversions.ToBoolean(KG.A.SettingsXml.GetElementsByTagName(AH.A(3218)).Item(0).InnerText);
					_ = null;
					C(pres, flag);
					break;
				}
			}
			else
			{
				flag = Conversions.ToBoolean(text);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return flag;
	}

	private static void D(Microsoft.Office.Interop.PowerPoint.Presentation A, bool B)
	{
		A.Tags.Add(Behavior.m_A, (B ? 1 : 0).ToString());
	}

	public static bool GetPresentationShowSubsections(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		bool flag = default(bool);
		try
		{
			string text = pres.Tags[Behavior.m_A];
			if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
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
					flag = Conversions.ToBoolean(KG.A.SettingsXml.GetElementsByTagName(AH.A(5483)).Item(0).InnerText);
					_ = null;
					D(pres, flag);
					break;
				}
			}
			else
			{
				flag = Conversions.ToBoolean(text);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return flag;
	}

	public static void ToggleAutoCollapse(bool pressed)
	{
		if (Licensing.AllowAgendaOperation())
		{
			Application application = NG.A.Application;
			Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
			application.StartNewUndoEntry();
			application = null;
			A(activePresentation, pressed);
			Update.RefreshAgendaStyle(activePresentation);
		}
	}

	public static void ToggleOmitDoubles(bool pressed)
	{
		if (!Licensing.AllowAgendaOperation())
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
			Application application = NG.A.Application;
			Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
			application.StartNewUndoEntry();
			application = null;
			C(activePresentation, pressed);
			Update.RefreshAgendaStyle(activePresentation);
			return;
		}
	}

	public static void ToggleSkipDoubles(bool pressed)
	{
		if (!Licensing.AllowAgendaOperation())
		{
			return;
		}
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
			Application application = NG.A.Application;
			Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
			application.StartNewUndoEntry();
			B(activePresentation, pressed);
			enumerator = activePresentation.Slides.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					Slide slide = (Slide)enumerator.Current;
					if (pressed)
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
						if (Helpers.GetSlideType(slide) != SlideType.Agenda)
						{
							continue;
						}
						if (slide.SlideIndex < activePresentation.Slides.Count)
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
							Slide slide2 = activePresentation.Slides[checked(slide.SlideIndex + 1)];
							if (Helpers.GetSlideType(slide2) == SlideType.Agenda)
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
								if (activePresentation.SectionProperties.Count > 0)
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
									if (Update.A(activePresentation.SectionProperties.Name(slide2.sectionIndex)))
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
										slide.SlideShowTransition.Hidden = MsoTriState.msoTrue;
									}
								}
							}
							slide2 = null;
						}
						else
						{
							slide.SlideShowTransition.Hidden = MsoTriState.msoFalse;
						}
					}
					else
					{
						slide.SlideShowTransition.Hidden = MsoTriState.msoFalse;
					}
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_0156;
					}
					continue;
					end_IL_0156:
					break;
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
			activePresentation = null;
			return;
		}
	}

	public static void ToggleShowSubsections(bool pressed)
	{
		if (!Licensing.AllowAgendaOperation())
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
			Application application = NG.A.Application;
			Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
			application.StartNewUndoEntry();
			_ = null;
			D(activePresentation, pressed);
			Update.A(A: true, activePresentation);
			activePresentation = null;
			return;
		}
	}

	public static bool FlysheetCollapsePressed()
	{
		return GetPresentationAutoCollapse(NG.A.Application.ActivePresentation);
	}

	public static bool FlysheetSkipPressed()
	{
		return GetPresentationSkipDoubles(NG.A.Application.ActivePresentation);
	}

	public static bool FlysheetOmitPressed()
	{
		return GetPresentationOmitDoubles(NG.A.Application.ActivePresentation);
	}

	public static bool FlysheetBehaviorEnabled()
	{
		bool result;
		try
		{
			result = GetPresentationFlysheetStyle(NG.A.Application.ActivePresentation) == FlySheetStyle.Agenda;
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

	public static bool FlysheetTopicPressed()
	{
		return A(FlySheetStyle.Topic);
	}

	public static bool FlysheetAgendaPressed()
	{
		return A(FlySheetStyle.Agenda);
	}

	private static bool A(FlySheetStyle A)
	{
		string text = NG.A.Application.ActivePresentation.Tags[F];
		if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
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
					return false;
				}
			}
		}
		return Conversions.ToInteger(text) == (int)A;
	}

	public static bool ShowSubsectionsPressed()
	{
		return GetPresentationShowSubsections(NG.A.Application.ActivePresentation);
	}
}
