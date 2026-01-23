using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Libraries.Manage.Publish;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Agenda;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.MasterShapes;

public sealed class Placeholders
{
	private static readonly string m_A = AH.A(120996);

	private static readonly string m_B = AH.A(120983);

	private static readonly string C = AH.A(120970);

	private static readonly string D = AH.A(120953);

	private static readonly string E = AH.A(120940);

	private static readonly string F = AH.A(120927);

	private static readonly string G = AH.A(120906);

	public static readonly string PLACEHOLDER_STAMP = AH.A(149107);

	public static readonly string PLACEHOLDER_SECTION = AH.A(149122);

	public static readonly string PLACEHOLDER_SUBSECTION = AH.A(149141);

	public static readonly string PLACEHOLDER_SEC_INDEX = AH.A(149166);

	public static readonly string PLACEHOLDER_SUBSEC_INDEX = AH.A(149185);

	public static readonly string PLACEHOLDER_TITLE = AH.A(149210);

	public static void Populate(Microsoft.Office.Interop.PowerPoint.Shape shp, Microsoft.Office.Interop.PowerPoint.Presentation pres, string strPlaceholder, DateTime dt, string strStamp, Slide sld = null)
	{
		checked
		{
			if (shp.HasTextFrame == MsoTriState.msoTrue)
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
						if (Placeholders.B(strPlaceholder))
						{
							while (true)
							{
								int sectionIndex;
								string text2;
								string text3;
								string A;
								string B;
								switch (3)
								{
								case 0:
									break;
								default:
									{
										strPlaceholder = strPlaceholder.Replace(Placeholders.m_A, Core.GetAuthor(pres));
										strPlaceholder = strPlaceholder.Replace(Placeholders.m_B, pres.FullName);
										strPlaceholder = strPlaceholder.Replace(C, pres.Name);
										strPlaceholder = strPlaceholder.Replace(D, pres.Path);
										strPlaceholder = strPlaceholder.Replace(PLACEHOLDER_STAMP, strStamp);
										if (strPlaceholder.Contains(E))
										{
											strPlaceholder = strPlaceholder.Replace(E, dt.ToShortDateString());
										}
										if (strPlaceholder.Contains(G))
										{
											strPlaceholder = strPlaceholder.Replace(G, dt.ToShortDateString() + AH.A(14625) + dt.ToShortTimeString());
										}
										if (strPlaceholder.Contains(F))
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
											strPlaceholder = strPlaceholder.Replace(F, dt.ToShortTimeString());
										}
										if (strPlaceholder.Contains(PLACEHOLDER_TITLE))
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
											int num = 1;
											do
											{
												if (Helpers.GetSlideType(pres.Slides[num]) == SlideType.Title)
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
													Microsoft.Office.Interop.PowerPoint.Shapes shapes = pres.Slides[num].Shapes;
													if (shapes.HasTitle == MsoTriState.msoTrue)
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
														try
														{
															strPlaceholder = strPlaceholder.Replace(PLACEHOLDER_TITLE, shapes.Title.TextFrame2.TextRange.Text);
														}
														catch (Exception ex)
														{
															ProjectData.SetProjectError(ex);
															Exception ex2 = ex;
															ProjectData.ClearProjectError();
														}
													}
													shapes = null;
													break;
												}
												num++;
											}
											while (num <= 2);
										}
										if (!strPlaceholder.Contains(PLACEHOLDER_SECTION))
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
											if (!strPlaceholder.Contains(PLACEHOLDER_SUBSECTION))
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
												if (!strPlaceholder.Contains(PLACEHOLDER_SEC_INDEX))
												{
													if (!strPlaceholder.Contains(PLACEHOLDER_SUBSEC_INDEX))
													{
														goto IL_04e7;
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
										}
										if (sld != null)
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
											if (pres.SectionProperties.Count > 1)
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
												sectionIndex = sld.sectionIndex;
												string text = pres.SectionProperties.Name(sectionIndex);
												text2 = "";
												text3 = "";
												A = "";
												B = "";
												if (!PowerPointAddIn1.Agenda.Update.A(text))
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
													text2 = text;
												}
												else
												{
													text3 = Strings.Mid(text, 2);
													int num2 = sectionIndex - 1;
													while (true)
													{
														if (num2 < 2)
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
															break;
														}
														if (!PowerPointAddIn1.Agenda.Update.A(pres.SectionProperties.Name(num2)))
														{
															text2 = pres.SectionProperties.Name(num2);
															break;
														}
														num2 += -1;
													}
												}
												if (sectionIndex == 1)
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
													if (Operators.CompareString(text2, AH.A(148954), TextCompare: false) != 0)
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
														if (Operators.CompareString(text2, AH.A(7090), TextCompare: false) != 0)
														{
															goto IL_0383;
														}
													}
													text2 = "";
													text3 = "";
												}
												goto IL_0383;
											}
										}
										goto IL_04e7;
									}
									IL_0383:
									if (text2.Length > 0)
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
										if (!strPlaceholder.Contains(PLACEHOLDER_SEC_INDEX))
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
											if (!strPlaceholder.Contains(PLACEHOLDER_SUBSEC_INDEX))
											{
												goto IL_03cf;
											}
										}
										Placeholders.A(ref A, ref B, sectionIndex, pres);
										goto IL_03cf;
									}
									strPlaceholder = "";
									goto IL_04e7;
									IL_03cf:
									if (text3.Length > 0)
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
										strPlaceholder = strPlaceholder.Replace(PLACEHOLDER_SECTION, text2);
										strPlaceholder = strPlaceholder.Replace(PLACEHOLDER_SUBSECTION, text3);
										if (A.Length > 0)
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
											strPlaceholder = strPlaceholder.Replace(PLACEHOLDER_SEC_INDEX, A);
										}
										if (B.Length > 0)
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
											strPlaceholder = strPlaceholder.Replace(PLACEHOLDER_SUBSEC_INDEX, B);
										}
									}
									else if (!strPlaceholder.Contains(PLACEHOLDER_SEC_INDEX))
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
										strPlaceholder = text2;
									}
									else
									{
										strPlaceholder = Strings.Left(strPlaceholder, strPlaceholder.IndexOf(PLACEHOLDER_SECTION) + PLACEHOLDER_SECTION.Length);
										strPlaceholder = strPlaceholder.Replace(PLACEHOLDER_SECTION, text2);
										if (A.Length > 0)
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
											strPlaceholder = strPlaceholder.Replace(PLACEHOLDER_SEC_INDEX, A);
										}
									}
									goto IL_04e7;
									IL_04e7:
									shp.TextFrame2.TextRange.Text = strPlaceholder;
									return;
								}
							}
						}
						return;
					}
				}
			}
			if (shp.Type != MsoShapeType.msoGroup)
			{
				return;
			}
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = shp.GroupItems.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Populate((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, pres, strPlaceholder, dt, strStamp, sld);
				}
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
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
		}
	}

	private static void A(ref string A, ref string B, int C, Microsoft.Office.Interop.PowerPoint.Presentation D)
	{
		List<string> B2 = null;
		List<string> C2 = null;
		int num = -1;
		int num2 = -1;
		bool flag = false;
		checked
		{
			for (int i = C; i >= 1; i += -1)
			{
				string text = D.SectionProperties.Name(i);
				if (PowerPointAddIn1.Agenda.Update.A(text))
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
					if (flag)
					{
						continue;
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
					num2++;
					continue;
				}
				if (Operators.CompareString(text, AH.A(148954), TextCompare: false) != 0)
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
					if (Operators.CompareString(text, AH.A(7090), TextCompare: false) != 0)
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
						num++;
					}
				}
				flag = true;
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				PowerPointAddIn1.Agenda.Update.A(Helpers.GetLayout(D, SlideType.Flysheet), ref B2, ref C2);
				if (B2 != null && num > -1)
				{
					A = Placeholders.A(B2[num]);
				}
				if (C2 != null)
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
					if (num2 > -1)
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
						B = Placeholders.A(C2[num2]);
					}
				}
				B2 = null;
				C2 = null;
				return;
			}
		}
	}

	private static string A(string A)
	{
		return A.Replace(AH.A(14417), "").Replace(AH.A(14255), "").Replace(AH.A(17795), "");
	}

	public static void Update()
	{
		Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
		string presentationStamp = Stamps.GetPresentationStamp(activePresentation);
		int num = 0;
		DateTime now = DateTime.Now;
		application.StartNewUndoEntry();
		checked
		{
			try
			{
				IEnumerator enumerator16 = default(IEnumerator);
				IEnumerator enumerator14 = default(IEnumerator);
				IEnumerator enumerator15 = default(IEnumerator);
				IEnumerator enumerator13 = default(IEnumerator);
				IEnumerator enumerator11 = default(IEnumerator);
				IEnumerator enumerator10 = default(IEnumerator);
				IEnumerator enumerator8 = default(IEnumerator);
				IEnumerator enumerator9 = default(IEnumerator);
				IEnumerator enumerator7 = default(IEnumerator);
				IEnumerator enumerator6 = default(IEnumerator);
				IEnumerator enumerator4 = default(IEnumerator);
				IEnumerator enumerator2 = default(IEnumerator);
				IEnumerator enumerator3 = default(IEnumerator);
				foreach (Microsoft.Office.Interop.PowerPoint.Shape shape3 in activePresentation.Designs[1].SlideMaster.Shapes)
				{
					Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape3;
					if (shape2.Visible == MsoTriState.msoFalse)
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
						if (Base.A(shape3))
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
							if (shape2.HasTextFrame == MsoTriState.msoTrue)
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
								if (shape2.TextFrame2.HasText == MsoTriState.msoTrue)
								{
									Behavior behavior = Base.A(shape2.Name);
									if (behavior != Behavior.SelectedSlides)
									{
										string text = shape2.TextFrame2.TextRange.Text;
										if (A(text))
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
											string c = shape2.Id.ToString();
											switch (behavior)
											{
											case Behavior.AllSlides:
												try
												{
													enumerator16 = activePresentation.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).GetEnumerator();
													while (enumerator16.MoveNext())
													{
														A((Slide)enumerator16.Current, text, c, activePresentation, now, presentationStamp);
														num++;
													}
													while (true)
													{
														switch (2)
														{
														case 0:
															break;
														default:
															goto end_IL_01b5;
														}
														continue;
														end_IL_01b5:
														break;
													}
												}
												finally
												{
													if (enumerator16 is IDisposable)
													{
														while (true)
														{
															switch (5)
															{
															case 0:
																continue;
															}
															(enumerator16 as IDisposable).Dispose();
															break;
														}
													}
												}
												break;
											case Behavior.AllLayouts:
												try
												{
													enumerator14 = activePresentation.Designs.GetEnumerator();
													while (enumerator14.MoveNext())
													{
														Design design5 = (Design)enumerator14.Current;
														try
														{
															enumerator15 = design5.SlideMaster.CustomLayouts.GetEnumerator();
															while (enumerator15.MoveNext())
															{
																A((CustomLayout)enumerator15.Current, text, c, activePresentation, now, presentationStamp);
																num++;
															}
															while (true)
															{
																switch (4)
																{
																case 0:
																	break;
																default:
																	goto end_IL_0249;
																}
																continue;
																end_IL_0249:
																break;
															}
														}
														finally
														{
															if (enumerator15 is IDisposable)
															{
																while (true)
																{
																	switch (6)
																	{
																	case 0:
																		continue;
																	}
																	(enumerator15 as IDisposable).Dispose();
																	break;
																}
															}
														}
													}
													while (true)
													{
														switch (6)
														{
														case 0:
															break;
														default:
															goto end_IL_0283;
														}
														continue;
														end_IL_0283:
														break;
													}
												}
												finally
												{
													if (enumerator14 is IDisposable)
													{
														while (true)
														{
															switch (4)
															{
															case 0:
																continue;
															}
															(enumerator14 as IDisposable).Dispose();
															break;
														}
													}
												}
												break;
											case Behavior.ContentSlides:
												try
												{
													enumerator13 = activePresentation.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).GetEnumerator();
													while (enumerator13.MoveNext())
													{
														Slide slide4 = (Slide)enumerator13.Current;
														if (Helpers.IsSpecialSlide(slide4))
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
														A(slide4, text, c, activePresentation, now, presentationStamp);
														num++;
													}
													while (true)
													{
														switch (2)
														{
														case 0:
															break;
														default:
															goto end_IL_0319;
														}
														continue;
														end_IL_0319:
														break;
													}
												}
												finally
												{
													if (enumerator13 is IDisposable)
													{
														while (true)
														{
															switch (2)
															{
															case 0:
																continue;
															}
															(enumerator13 as IDisposable).Dispose();
															break;
														}
													}
												}
												break;
											case Behavior.ContentLayouts:
												try
												{
													enumerator11 = activePresentation.Designs.GetEnumerator();
													while (enumerator11.MoveNext())
													{
														Design design4 = (Design)enumerator11.Current;
														foreach (CustomLayout customLayout4 in design4.SlideMaster.CustomLayouts)
														{
															if (Helpers.IsSpecialLayout(customLayout4))
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
															A(customLayout4, text, c, activePresentation, now, presentationStamp);
															num++;
														}
													}
													while (true)
													{
														switch (5)
														{
														case 0:
															break;
														default:
															goto end_IL_03ec;
														}
														continue;
														end_IL_03ec:
														break;
													}
												}
												finally
												{
													if (enumerator11 is IDisposable)
													{
														while (true)
														{
															switch (2)
															{
															case 0:
																continue;
															}
															(enumerator11 as IDisposable).Dispose();
															break;
														}
													}
												}
												break;
											case Behavior.SlidesShowingBackgroundGraphics:
												try
												{
													enumerator10 = activePresentation.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).GetEnumerator();
													while (enumerator10.MoveNext())
													{
														Slide slide3 = (Slide)enumerator10.Current;
														if (slide3.CustomLayout.DisplayMasterShapes != MsoTriState.msoTrue)
														{
															continue;
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
														A(slide3, text, c, activePresentation, now, presentationStamp);
														num++;
													}
												}
												finally
												{
													if (enumerator10 is IDisposable)
													{
														while (true)
														{
															switch (1)
															{
															case 0:
																continue;
															}
															(enumerator10 as IDisposable).Dispose();
															break;
														}
													}
												}
												break;
											case Behavior.LayoutsShowingBackgroundGraphics:
												try
												{
													enumerator8 = activePresentation.Designs.GetEnumerator();
													while (enumerator8.MoveNext())
													{
														Design design3 = (Design)enumerator8.Current;
														{
															enumerator9 = design3.SlideMaster.CustomLayouts.GetEnumerator();
															try
															{
																while (enumerator9.MoveNext())
																{
																	CustomLayout customLayout3 = (CustomLayout)enumerator9.Current;
																	if (customLayout3.DisplayMasterShapes != MsoTriState.msoTrue)
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
																		break;
																	}
																	A(customLayout3, text, c, activePresentation, now, presentationStamp);
																	num++;
																}
																while (true)
																{
																	switch (3)
																	{
																	case 0:
																		break;
																	default:
																		goto end_IL_0532;
																	}
																	continue;
																	end_IL_0532:
																	break;
																}
															}
															finally
															{
																IDisposable disposable = enumerator9 as IDisposable;
																if (disposable != null)
																{
																	disposable.Dispose();
																}
															}
														}
													}
													while (true)
													{
														switch (3)
														{
														case 0:
															break;
														default:
															goto end_IL_0562;
														}
														continue;
														end_IL_0562:
														break;
													}
												}
												finally
												{
													if (enumerator8 is IDisposable)
													{
														while (true)
														{
															switch (4)
															{
															case 0:
																continue;
															}
															(enumerator8 as IDisposable).Dispose();
															break;
														}
													}
												}
												break;
											case Behavior.DynamicSlides:
												try
												{
													enumerator7 = activePresentation.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).GetEnumerator();
													while (enumerator7.MoveNext())
													{
														Slide slide2 = (Slide)enumerator7.Current;
														SlideType slideType = Helpers.GetSlideType(slide2);
														unchecked
														{
															if ((uint)(slideType - 4) <= 1u)
															{
																continue;
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
															if ((uint)(slideType - 9) <= 1u)
															{
																continue;
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
															A(slide2, text, c, activePresentation, now, presentationStamp);
														}
														num++;
													}
													while (true)
													{
														switch (5)
														{
														case 0:
															break;
														default:
															goto end_IL_0611;
														}
														continue;
														end_IL_0611:
														break;
													}
												}
												finally
												{
													if (enumerator7 is IDisposable)
													{
														while (true)
														{
															switch (4)
															{
															case 0:
																continue;
															}
															(enumerator7 as IDisposable).Dispose();
															break;
														}
													}
												}
												break;
											case Behavior.DynamicLayouts:
												foreach (Design design6 in activePresentation.Designs)
												{
													try
													{
														enumerator6 = design6.SlideMaster.CustomLayouts.GetEnumerator();
														while (enumerator6.MoveNext())
														{
															CustomLayout customLayout2 = (CustomLayout)enumerator6.Current;
															SlideType layoutType = Helpers.GetLayoutType(customLayout2);
															if (unchecked((uint)(layoutType - 4) > 1u && (uint)(layoutType - 9) > 1u))
															{
																A(customLayout2, text, c, activePresentation, now, presentationStamp);
																num++;
															}
														}
														while (true)
														{
															switch (4)
															{
															case 0:
																break;
															default:
																goto end_IL_06c7;
															}
															continue;
															end_IL_06c7:
															break;
														}
													}
													finally
													{
														if (enumerator6 is IDisposable)
														{
															while (true)
															{
																switch (5)
																{
																case 0:
																	continue;
																}
																(enumerator6 as IDisposable).Dispose();
																break;
															}
														}
													}
												}
												break;
											case Behavior.SpecialSlides:
												try
												{
													enumerator4 = activePresentation.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).GetEnumerator();
													while (enumerator4.MoveNext())
													{
														Slide slide = (Slide)enumerator4.Current;
														if (Helpers.IsSpecialSlide(slide))
														{
															A(slide, text, c, activePresentation, now, presentationStamp);
															num++;
														}
													}
												}
												finally
												{
													if (enumerator4 is IDisposable)
													{
														while (true)
														{
															switch (5)
															{
															case 0:
																continue;
															}
															(enumerator4 as IDisposable).Dispose();
															break;
														}
													}
												}
												break;
											case Behavior.SpecialLayouts:
												{
													enumerator2 = activePresentation.Designs.GetEnumerator();
													try
													{
														while (enumerator2.MoveNext())
														{
															Design design = (Design)enumerator2.Current;
															try
															{
																enumerator3 = design.SlideMaster.CustomLayouts.GetEnumerator();
																while (enumerator3.MoveNext())
																{
																	CustomLayout customLayout = (CustomLayout)enumerator3.Current;
																	if (!Helpers.IsSpecialLayout(customLayout))
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
																		break;
																	}
																	A(customLayout, text, c, activePresentation, now, presentationStamp);
																	num++;
																}
															}
															finally
															{
																if (enumerator3 is IDisposable)
																{
																	while (true)
																	{
																		switch (7)
																		{
																		case 0:
																			continue;
																		}
																		(enumerator3 as IDisposable).Dispose();
																		break;
																	}
																}
															}
														}
														while (true)
														{
															switch (5)
															{
															case 0:
																break;
															default:
																goto end_IL_0852;
															}
															continue;
															end_IL_0852:
															break;
														}
													}
													finally
													{
														IDisposable disposable2 = enumerator2 as IDisposable;
														if (disposable2 != null)
														{
															disposable2.Dispose();
														}
													}
												}
												break;
											}
										}
									}
								}
							}
						}
					}
					shape2 = null;
				}
				Forms.SuccessMessage(AH.A(148985));
				Base.A(AH.A(149068));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception a = ex;
				Base.A(a);
				ProjectData.ClearProjectError();
			}
			activePresentation = null;
		}
	}

	private static void A(Slide A, string B, string C, Microsoft.Office.Interop.PowerPoint.Presentation D, DateTime E, string F)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Shapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				if (!Base.A(shape, C))
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Populate(shape, D, B, E, F, A);
			}
			while (true)
			{
				switch (3)
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
	}

	private static void A(CustomLayout A, string B, string C, Microsoft.Office.Interop.PowerPoint.Presentation D, DateTime E, string F)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Shapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				if (!Base.A(shape, C))
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
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Populate(shape, D, B, E, F);
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	private static bool A(string A)
	{
		if (A.Length > 5 && A.Contains(AH.A(15135)))
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
					return A.Contains(AH.A(15138));
				}
			}
		}
		return false;
	}

	private static bool B(string A)
	{
		if (Placeholders.A(A))
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
					if (A.Contains(Placeholders.m_A))
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								return true;
							}
						}
					}
					if (A.Contains(E))
					{
						return true;
					}
					if (A.Contains(G))
					{
						return true;
					}
					if (A.Contains(C))
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								return true;
							}
						}
					}
					if (A.Contains(D))
					{
						return true;
					}
					if (A.Contains(Placeholders.m_B))
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								return true;
							}
						}
					}
					if (A.Contains(PLACEHOLDER_SECTION))
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								return true;
							}
						}
					}
					if (A.Contains(PLACEHOLDER_STAMP))
					{
						return true;
					}
					if (A.Contains(PLACEHOLDER_SUBSECTION))
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								return true;
							}
						}
					}
					if (A.Contains(F))
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								return true;
							}
						}
					}
					if (A.Contains(PLACEHOLDER_TITLE))
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								return true;
							}
						}
					}
					if (A.Contains(PLACEHOLDER_SEC_INDEX))
					{
						return true;
					}
					if (A.Contains(PLACEHOLDER_SUBSEC_INDEX))
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								return true;
							}
						}
					}
					return false;
				}
			}
		}
		return false;
	}
}
