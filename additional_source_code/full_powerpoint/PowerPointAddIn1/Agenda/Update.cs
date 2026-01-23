using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Explorer;
using PowerPointAddIn1.MasterShapes;
using PowerPointAddIn1.Shapes;
using PowerPointAddIn1.Slides;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1.Agenda;

public sealed class Update
{
	private struct QB
	{
		public Slide A;

		public string A;

		public int A;

		public int B;

		public int C;
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<string, string> A;

		public static Func<string, string> B;

		public static Func<string, string> C;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal string A(string A)
		{
			return A + AH.A(14417);
		}

		[SpecialName]
		internal string B(string A)
		{
			return A + AH.A(14255);
		}

		[SpecialName]
		internal string C(string A)
		{
			return AH.A(17795) + A + AH.A(14255);
		}
	}

	[CompilerGenerated]
	internal sealed class RB
	{
		public Slide A;

		public Func<QB, bool> A;

		public RB(RB A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(QB A)
		{
			if (A.A != null)
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
						return A.A.SlideIndex == this.A.SlideIndex;
					}
				}
			}
			return false;
		}
	}

	[CompilerGenerated]
	internal sealed class SB
	{
		public Slide A;

		public SB(SB A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal bool A(QB A)
		{
			return A.A.SlideIndex == this.A.SlideIndex;
		}
	}

	internal static void A(bool A, Microsoft.Office.Interop.PowerPoint.Presentation B = null)
	{
		if (B == null)
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
			B = NG.A.Application.ActivePresentation;
		}
		Events.Disable(B.Application);
		if (Behavior.GetPresentationFlysheetStyle(B) == FlySheetStyle.Topic)
		{
			RefreshTopicStyle(B);
		}
		else
		{
			RefreshAgendaStyle(B);
		}
		SectionTitles.Update(B, A);
		B = null;
	}

	public static void RefreshAgendaStyle(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		List<string> B = null;
		List<string> C = null;
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		Dictionary<int, int> dictionary2 = new Dictionary<int, int>();
		Dictionary<int, float> D = new Dictionary<int, float>();
		float num = 0f;
		float C2 = 0f;
		float B2 = 0f;
		float num2 = 0f;
		bool C3 = false;
		Slide slide = null;
		Slide slide2 = null;
		CustomLayout customLayout = null;
		if (pres == null)
		{
			return;
		}
		checked
		{
			bool presentationShowSubsections = default(bool);
			IEnumerator enumerator2 = default(IEnumerator);
			IEnumerator enumerator3 = default(IEnumerator);
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				slide = TableOfContents.Slide(pres);
				if (slide != null)
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
					customLayout = slide.CustomLayout;
				}
				else
				{
					customLayout = PowerPointAddIn1.Slides.Helpers.GetLayout(pres, SlideType.Agenda);
					if (customLayout == null)
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
						customLayout = PowerPointAddIn1.Slides.Helpers.GetLayout(pres, SlideType.TableOfContents);
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
							PowerPointAddIn1.Slides.Helpers.DesignateLayoutAsType(customLayout, SlideType.Agenda);
						}
					}
				}
				List<QB> list;
				if (customLayout != null)
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
					list = A(pres, customLayout, ref C3);
					A(pres, list, SlideType.Agenda, slide);
					if (C3)
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
							presentationShowSubsections = Behavior.GetPresentationShowSubsections(pres);
						}
					}
					int num3 = list.Count - 1;
					int num4 = num3;
					for (int i = 0; i <= num4; i++)
					{
						QB value = list[i];
						value.C = A(value.A);
						if (value.C == 0)
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
							if (value.A == null)
							{
								int num5 = value.B + 1;
								int num6 = num3;
								for (int j = num5; j <= num6; j++)
								{
									int num7 = pres.SectionProperties.FirstSlide(j);
									if (num7 > 0)
									{
										value.C = A(pres.Slides[num7]);
										break;
									}
								}
							}
						}
						list[i] = value;
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
					if (slide != null)
					{
						A(slide);
					}
					if (list.Any())
					{
						Dictionary<int, float> dictionary3 = D;
						dictionary3.Add(1, 0f);
						dictionary3.Add(2, 0f);
						_ = null;
						A(customLayout, ref B2, ref C2, ref D);
						A(PowerPointAddIn1.Slides.Helpers.GetLayout(pres, SlideType.Flysheet), ref B, ref C);
						num2 = B2;
						Dictionary<int, int> dictionary4 = dictionary2;
						dictionary4.Add(1, 0);
						dictionary4.Add(2, 0);
						_ = null;
						if (B != null)
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
							if (C != null)
							{
								bool presentationAutoCollapse = Behavior.GetPresentationAutoCollapse(pres);
								bool presentationSkipDoubles = Behavior.GetPresentationSkipDoubles(pres);
								int num8 = num3;
								for (int k = 0; k <= num8; k++)
								{
									QB qB = list[k];
									int num9;
									SlideShowTransition slideShowTransition;
									if (qB.A != null)
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
										num9 = 0;
										num = B2;
										Dictionary<int, int> dictionary5 = dictionary;
										dictionary5.Clear();
										dictionary5.Add(1, 0);
										dictionary5.Add(2, 0);
										_ = null;
										A(qB.A);
										if (qB.A == 1)
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
											slide2 = qB.A;
											slideShowTransition = qB.A.SlideShowTransition;
											if (presentationSkipDoubles)
											{
												if (k < num3)
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
													if (list[k].A != null)
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
														if (pres.Slides[list[k].A.SlideIndex + 1] == list[k + 1].A)
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
															if (list[k + 1].A == 2)
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
																slideShowTransition.Hidden = MsoTriState.msoTrue;
																goto IL_039e;
															}
														}
													}
												}
												slideShowTransition.Hidden = MsoTriState.msoFalse;
											}
											else
											{
												slideShowTransition.Hidden = MsoTriState.msoFalse;
											}
											goto IL_039e;
										}
										goto IL_03a2;
									}
									goto IL_07ac;
									IL_03a2:
									int a2;
									string text;
									using (List<QB>.Enumerator enumerator = list.GetEnumerator())
									{
										while (enumerator.MoveNext())
										{
											QB current = enumerator.Current;
											string a = current.A;
											a2 = current.A;
											dictionary[a2]++;
											if (a2 == 1)
											{
												text = B[dictionary[a2] - 1];
												dictionary[2] = 0;
											}
											else
											{
												if (presentationAutoCollapse)
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
													if (slide2 != null)
													{
														if (current.B != qB.B)
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
													}
												}
												text = C[dictionary[a2] - 1];
											}
											num9++;
											Microsoft.Office.Interop.PowerPoint.Shape shape = A(qB.A, customLayout, a2, qB.A == current.A, ref num, D[a2], C2);
											try
											{
												enumerator2 = shape.GroupItems.GetEnumerator();
												while (enumerator2.MoveNext())
												{
													Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
													if (shape2.HasTextFrame == MsoTriState.msoFalse)
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
													TextRange2 textRange = shape2.TextFrame2.TextRange;
													string text2 = textRange.Text.ToUpper();
													uint num10 = YG.A(text2);
													if (num10 <= 211792314)
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
														if (num10 != 7274958)
														{
															if (num10 != 149826459)
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
																if (num10 == 211792314 && Operators.CompareString(text2, AH.A(7299), TextCompare: false) == 0)
																{
																	goto IL_06a4;
																}
															}
															else
															{
																if (Operators.CompareString(text2, AH.A(7342), TextCompare: false) == 0)
																{
																	goto IL_06a4;
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
														else if (Operators.CompareString(text2, AH.A(7312), TextCompare: false) == 0)
														{
															goto IL_06a4;
														}
													}
													else if (num10 <= 2063731641)
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
														if (num10 != 2050578930)
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
															if (num10 != 2063731641)
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
															else if (Operators.CompareString(text2, AH.A(7365), TextCompare: false) != 0)
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
															else if (current.C > 0)
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
																if (current.A == null)
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
																	if (qB.B == current.B)
																	{
																		goto IL_070d;
																	}
																}
																if (!presentationAutoCollapse)
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
																	if (current.A == null)
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
																		goto IL_070d;
																	}
																}
																textRange.Text = current.C.ToString();
															}
															else
															{
																textRange.Text = "";
															}
														}
														else if (Operators.CompareString(text2, AH.A(7277), TextCompare: false) == 0)
														{
															textRange.Text = text;
														}
													}
													else if (num10 != 2381896112u)
													{
														if (num10 != 3278918548u)
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
															if (Operators.CompareString(text2, AH.A(7327), TextCompare: false) == 0)
															{
																goto IL_06a4;
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
														if (Operators.CompareString(text2, AH.A(7284), TextCompare: false) == 0)
														{
															goto IL_06a4;
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
													goto IL_073c;
													IL_073c:
													textRange = null;
													continue;
													IL_070d:
													textRange.Text = "";
													goto IL_073c;
													IL_06a4:
													textRange.Text = a;
													goto IL_073c;
												}
												while (true)
												{
													switch (6)
													{
													case 0:
														break;
													default:
														goto end_IL_074d;
													}
													continue;
													end_IL_074d:
													break;
												}
											}
											finally
											{
												if (enumerator2 is IDisposable)
												{
													while (true)
													{
														switch (3)
														{
														case 0:
															continue;
														}
														(enumerator2 as IDisposable).Dispose();
														break;
													}
												}
											}
											shape.Visible = MsoTriState.msoTrue;
											shape = null;
										}
										while (true)
										{
											switch (5)
											{
											case 0:
												break;
											default:
												goto end_IL_0792;
											}
											continue;
											end_IL_0792:
											break;
										}
									}
									goto IL_07ac;
									IL_039e:
									slideShowTransition = null;
									goto IL_03a2;
									IL_07ac:
									if (slide == null)
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
									if (qB.A != 1)
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
										if (!presentationShowSubsections)
										{
											continue;
										}
									}
									a2 = qB.A;
									dictionary2[a2]++;
									if (a2 == 1)
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
										text = B[dictionary2[a2] - 1];
										dictionary2[2] = 0;
									}
									else
									{
										text = C[dictionary2[a2] - 1];
									}
									Microsoft.Office.Interop.PowerPoint.Shape shape3 = A(slide, customLayout, a2, D: false, ref num2, D[a2], C2);
									try
									{
										enumerator3 = shape3.GroupItems.GetEnumerator();
										while (enumerator3.MoveNext())
										{
											Microsoft.Office.Interop.PowerPoint.Shape shape4 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator3.Current;
											if (shape4.HasTextFrame == MsoTriState.msoFalse)
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
											TextRange2 textRange2 = shape4.TextFrame2.TextRange;
											string text3 = textRange2.Text.ToUpper();
											uint num10 = YG.A(text3);
											if (num10 <= 211792314)
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
												if (num10 != 7274958)
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
													if (num10 != 149826459)
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
														if (num10 != 211792314)
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
															if (Operators.CompareString(text3, AH.A(7299), TextCompare: false) == 0)
															{
																goto IL_0a92;
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
														if (Operators.CompareString(text3, AH.A(7342), TextCompare: false) == 0)
														{
															goto IL_0a92;
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
												else
												{
													if (Operators.CompareString(text3, AH.A(7312), TextCompare: false) == 0)
													{
														goto IL_0a92;
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
											else if (num10 <= 2063731641)
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
												if (num10 != 2050578930)
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
													if (num10 == 2063731641)
													{
														if (Operators.CompareString(text3, AH.A(7365), TextCompare: false) != 0)
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
														else if (qB.C > 0)
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
															textRange2.Text = qB.C.ToString();
														}
														else
														{
															textRange2.Text = "";
														}
													}
												}
												else if (Operators.CompareString(text3, AH.A(7277), TextCompare: false) != 0)
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
													textRange2.Text = text;
												}
											}
											else if (num10 != 2381896112u)
											{
												if (num10 != 3278918548u)
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
													if (Operators.CompareString(text3, AH.A(7327), TextCompare: false) == 0)
													{
														goto IL_0a92;
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
											else if (Operators.CompareString(text3, AH.A(7284), TextCompare: false) == 0)
											{
												goto IL_0a92;
											}
											goto IL_0ad9;
											IL_0a92:
											textRange2.Text = qB.A;
											goto IL_0ad9;
											IL_0ad9:
											textRange2 = null;
										}
										while (true)
										{
											switch (5)
											{
											case 0:
												break;
											default:
												goto end_IL_0ae8;
											}
											continue;
											end_IL_0ae8:
											break;
										}
									}
									finally
									{
										if (enumerator3 is IDisposable)
										{
											while (true)
											{
												switch (6)
												{
												case 0:
													continue;
												}
												(enumerator3 as IDisposable).Dispose();
												break;
											}
										}
									}
									shape3.Visible = MsoTriState.msoTrue;
									shape3 = null;
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
								goto IL_0b4b;
							}
						}
						Forms.ErrorMessage(AH.A(7372));
						goto IL_0b4b;
					}
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
						A(pres, slide);
					}
				}
				else
				{
					Forms.WarningMessage(AH.A(7632));
				}
				goto IL_0b7a;
				IL_0b7a:
				list = null;
				customLayout = null;
				slide = null;
				dictionary = null;
				dictionary2 = null;
				D = null;
				Events.RefreshPresentation(pres);
				return;
				IL_0b4b:
				B = null;
				C = null;
				goto IL_0b7a;
			}
		}
	}

	private static List<QB> A(Microsoft.Office.Interop.PowerPoint.Presentation A, CustomLayout B, ref bool C)
	{
		List<QB> list = new List<QB>();
		bool presentationOmitDoubles = Behavior.GetPresentationOmitDoubles(A);
		SectionProperties sectionProperties = A.SectionProperties;
		int count = sectionProperties.Count;
		int num = count;
		checked
		{
			Slide slide = default(Slide);
			int num4 = default(int);
			int b = default(int);
			for (int num2 = 2; num2 <= num; num2++)
			{
				int num3 = sectionProperties.FirstSlide(num2);
				if (num3 != -1)
				{
					if (PowerPointAddIn1.Slides.Helpers.GetSlideType(A.Slides[num3]) == SlideType.Agenda)
					{
						if (presentationOmitDoubles)
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
							if (num2 < count && Update.A(A.SectionProperties.Name(num2 + 1)))
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
								if (sectionProperties.SlidesCount(num2) == 1)
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
									if (!Update.A(sectionProperties.Name(num2)) && PowerPointAddIn1.Slides.Helpers.GetSlideType(A.Slides[num3]) == SlideType.Agenda)
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
										A.Slides[num3].Delete();
										slide = null;
										goto IL_0237;
									}
								}
							}
						}
						slide = A.Slides[num3];
						goto IL_0237;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
				}
				if (num3 == -1)
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
					if (num2 < count)
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
						if (presentationOmitDoubles)
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
							if (Update.A(A.SectionProperties.Name(num2 + 1)))
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
								slide = null;
								num4 = 0;
								goto IL_0143;
							}
						}
						bool flag = false;
						int num5 = num2;
						int num6 = count;
						int num7 = num5;
						while (true)
						{
							if (num7 <= num6)
							{
								if (sectionProperties.SlidesCount(num7) > 0)
								{
									num4 = sectionProperties.FirstSlide(num7) - 1;
									flag = true;
									break;
								}
								num7++;
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
							break;
						}
						if (!flag)
						{
							num4 = A.Slides.Count;
						}
					}
					else
					{
						num4 = A.Slides.Count;
					}
				}
				else
				{
					num4 = num3 - 1;
				}
				goto IL_0143;
				IL_0143:
				if (num4 > 0)
				{
					slide = A.Slides.AddSlide(num4 + 1, B);
					Create.SkipPresentationNewSlideEvent(slide);
					slide.MoveToSectionStart(num2);
				}
				goto IL_0237;
				IL_0237:
				QB item = new QB
				{
					A = slide
				};
				if (!Update.A(sectionProperties.Name(num2)))
				{
					item.A = 1;
					item.A = sectionProperties.Name(num2);
				}
				else
				{
					item.A = 2;
					item.A = Strings.Mid(sectionProperties.Name(num2), 2);
					C = true;
				}
				if (item.A == 1)
				{
					b = num2;
				}
				item.B = b;
				list.Add(item);
				slide = null;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				sectionProperties = null;
				return list;
			}
		}
	}

	public static void RefreshTopicStyle(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		List<string> B = null;
		List<string> C = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		bool flag = false;
		int num = 0;
		float B2 = 0f;
		float C2 = 0f;
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		Dictionary<int, float> D = new Dictionary<int, float>();
		bool C3 = false;
		CustomLayout customLayout = null;
		CustomLayout customLayout2 = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = null;
		Slide slide = null;
		if (pres == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		bool presentationShowSubsections = default(bool);
		IEnumerator enumerator3 = default(IEnumerator);
		SB sB = default(SB);
		IEnumerator enumerator4 = default(IEnumerator);
		IEnumerator enumerator5 = default(IEnumerator);
		IEnumerator enumerator6 = default(IEnumerator);
		IEnumerator enumerator7 = default(IEnumerator);
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (slide == null)
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
				slide = TableOfContents.Slide(pres);
				if (slide != null)
				{
					customLayout = slide.CustomLayout;
				}
			}
			else
			{
				customLayout = slide.CustomLayout;
			}
			try
			{
				enumerator = pres.Designs.GetEnumerator();
				while (true)
				{
					if (enumerator.MoveNext())
					{
						Design design = (Design)enumerator.Current;
						try
						{
							enumerator2 = design.SlideMaster.CustomLayouts.GetEnumerator();
							while (true)
							{
								if (enumerator2.MoveNext())
								{
									CustomLayout customLayout3 = (CustomLayout)enumerator2.Current;
									shape2 = PowerPointAddIn1.Slides.Helpers.GetSlideTypeShape(customLayout3.Shapes);
									if (shape2 == null)
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
									try
									{
										SlideType slideType = (SlideType)Conversions.ToInteger(shape2.Tags[PowerPointAddIn1.Slides.Helpers.TAG_SLIDE_TYPE].ToString());
										if (slideType != SlideType.TableOfContents)
										{
											while (true)
											{
												switch (6)
												{
												case 0:
													continue;
												}
												if (slideType != SlideType.Flysheet)
												{
													while (true)
													{
														switch (6)
														{
														case 0:
															break;
														default:
															goto end_IL_012b;
														}
														continue;
														end_IL_012b:
														break;
													}
													break;
												}
												customLayout2 = customLayout3;
												if (customLayout == null)
												{
													break;
												}
												while (true)
												{
													switch (1)
													{
													case 0:
														break;
													default:
														goto end_IL_0153;
													}
													continue;
													end_IL_0153:
													break;
												}
												goto end_IL_0170;
											}
											continue;
										}
										customLayout = customLayout3;
										if (customLayout2 == null)
										{
											continue;
										}
										while (true)
										{
											switch (3)
											{
											case 0:
												break;
											default:
												goto end_IL_013f;
											}
											continue;
											end_IL_013f:
											break;
										}
										break;
									}
									catch (Exception ex)
									{
										ProjectData.SetProjectError(ex);
										Exception ex2 = ex;
										ProjectData.ClearProjectError();
									}
									continue;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_017e;
									}
									continue;
									end_IL_017e:
									break;
								}
								break;
								continue;
								end_IL_0170:
								break;
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									(enumerator2 as IDisposable).Dispose();
									break;
								}
							}
						}
						shape2 = null;
						if (customLayout == null)
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
						if (customLayout2 == null)
						{
							continue;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_01bf;
							}
							continue;
							end_IL_01bf:
							break;
						}
						break;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_01d7;
						}
						continue;
						end_IL_01d7:
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
			checked
			{
				if (customLayout != null)
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
					List<QB> list = Update.B(pres, customLayout2, ref C3);
					if (slide != null)
					{
						A(slide);
					}
					if (list.Any())
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
						A(pres, list, SlideType.Flysheet, slide);
						if (C3)
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
							if (slide != null)
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
								presentationShowSubsections = Behavior.GetPresentationShowSubsections(pres);
							}
						}
						Dictionary<int, int> dictionary2 = dictionary;
						dictionary2.Add(1, 0);
						dictionary2.Add(2, 0);
						dictionary2.Add(3, 0);
						_ = null;
						Dictionary<int, float> dictionary3 = D;
						dictionary3.Add(1, 0f);
						dictionary3.Add(2, 0f);
						dictionary3.Add(3, 0f);
						_ = null;
						A(customLayout, ref B2, ref C2, ref D);
						A(customLayout2, ref B, ref C);
						if (B != null)
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
							if (C != null)
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
								try
								{
									enumerator3 = pres.Slides.GetEnumerator();
									while (enumerator3.MoveNext())
									{
										Slide slide2 = (Slide)enumerator3.Current;
										if (PowerPointAddIn1.Slides.Helpers.GetSlideType(slide2) == SlideType.Flysheet)
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
											sB = new SB(sB);
											sB.A = slide2;
											QB qB = list.Where(sB.A).ToList()[0];
											int a = qB.A;
											string a2 = qB.A;
											sB.A = null;
											Microsoft.Office.Interop.PowerPoint.Shape flysheetTopicShape = Flysheets.GetFlysheetTopicShape(slide2);
											if (flysheetTopicShape == null)
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
											dictionary[a]++;
											string d;
											if (a == 1)
											{
												d = B[dictionary[a] - 1];
												dictionary[2] = 0;
												dictionary[3] = 0;
											}
											else
											{
												d = C[dictionary[a] - 1];
											}
											TextRange2 textRange = flysheetTopicShape.TextFrame2.TextRange;
											textRange.Text = a2;
											ParagraphFormat2 paragraphFormat = textRange.ParagraphFormat;
											paragraphFormat.Bullet.StartValue = 1;
											paragraphFormat.IndentLevel = a;
											paragraphFormat.Bullet.StartValue = dictionary[a];
											_ = null;
											_ = null;
											num++;
											if (slide == null)
											{
												continue;
											}
											if (a != 1)
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
												if (!presentationShowSubsections)
												{
													continue;
												}
											}
											A(customLayout, slide, flysheetTopicShape, d, A(slide2), num, a, ref B2, D[a], C2);
											continue;
										}
										if (PowerPointAddIn1.Slides.Helpers.IsSpecialSlide(slide2))
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
										if (slide == null)
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
										try
										{
											enumerator4 = slide2.Shapes.GetEnumerator();
											while (enumerator4.MoveNext())
											{
												Microsoft.Office.Interop.PowerPoint.Shape shape3 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator4.Current;
												try
												{
													if (shape3.HasTextFrame != MsoTriState.msoTrue)
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
														if (shape3.TextFrame2.HasText != MsoTriState.msoTrue)
														{
															break;
														}
														while (true)
														{
															switch (2)
															{
															case 0:
																continue;
															}
															int a = A(shape3);
															if (a <= 0)
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
																num++;
																dictionary[a]++;
																string d;
																if (a != 1)
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
																	if (a != 2)
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
																		d = "";
																	}
																	else
																	{
																		d = C[dictionary[a] - 1];
																	}
																}
																else
																{
																	d = B[dictionary[a] - 1];
																	dictionary[2] = 0;
																	dictionary[3] = 0;
																}
																A(customLayout, slide, shape3, d, A(slide2), num, a, ref B2, D[a], C2);
																break;
															}
															break;
														}
														break;
													}
													continue;
												}
												catch (Exception ex3)
												{
													ProjectData.SetProjectError(ex3);
													Exception ex4 = ex3;
													ProjectData.ClearProjectError();
												}
											}
											while (true)
											{
												switch (3)
												{
												case 0:
													break;
												default:
													goto end_IL_05fe;
												}
												continue;
												end_IL_05fe:
												break;
											}
										}
										finally
										{
											if (enumerator4 is IDisposable)
											{
												while (true)
												{
													switch (7)
													{
													case 0:
														continue;
													}
													(enumerator4 as IDisposable).Dispose();
													break;
												}
											}
										}
									}
								}
								finally
								{
									if (enumerator3 is IDisposable)
									{
										while (true)
										{
											switch (1)
											{
											case 0:
												continue;
											}
											(enumerator3 as IDisposable).Dispose();
											break;
										}
									}
								}
								goto IL_066b;
							}
						}
						Forms.ErrorMessage(AH.A(7372));
						goto IL_066b;
					}
					bool flag2 = false;
					try
					{
						flag2 = customLayout.Shapes[Constants.AGENDA_TITLE_LEVEL_1] == null;
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						flag2 = true;
						ProjectData.ClearProjectError();
					}
					if (!flag2)
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
						if (slide != null)
						{
							A(pres, slide);
						}
					}
					else if (slide != null)
					{
						try
						{
							enumerator5 = slide.Shapes.GetEnumerator();
							while (true)
							{
								if (enumerator5.MoveNext())
								{
									Microsoft.Office.Interop.PowerPoint.Shape shape4 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator5.Current;
									if (Operators.CompareString(PowerPointAddIn1.Shapes.Helpers.GetShapeType(shape4), Constants.SHAPE_TYPE_TOC, TextCompare: false) == 0)
									{
										if (shape4.HasTextFrame == MsoTriState.msoTrue)
										{
											shape = shape4;
											shape.TextFrame.TextRange.Text = "";
										}
										break;
									}
									continue;
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_0733;
									}
									continue;
									end_IL_0733:
									break;
								}
								break;
							}
						}
						finally
						{
							if (enumerator5 is IDisposable)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									(enumerator5 as IDisposable).Dispose();
									break;
								}
							}
						}
						if (shape != null)
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
								enumerator6 = pres.Slides.GetEnumerator();
								while (enumerator6.MoveNext())
								{
									Slide slide3 = (Slide)enumerator6.Current;
									if (PowerPointAddIn1.Slides.Helpers.GetSlideType(slide3) != SlideType.Flysheet)
									{
										continue;
									}
									try
									{
										enumerator7 = slide3.Shapes.GetEnumerator();
										while (enumerator7.MoveNext())
										{
											Microsoft.Office.Interop.PowerPoint.Shape shape5 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator7.Current;
											try
											{
												Microsoft.Office.Interop.PowerPoint.Shape shape6 = shape5;
												if (shape6.HasTextFrame == MsoTriState.msoTrue)
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
													if (shape6.TextFrame2.HasText == MsoTriState.msoTrue)
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
														string shapeType = PowerPointAddIn1.Shapes.Helpers.GetShapeType(shape5);
														if (Operators.CompareString(shapeType, Constants.A, TextCompare: false) == 0)
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
															num++;
															A(shape, shape5, num, 1);
														}
														else if (Operators.CompareString(shapeType, Constants.B, TextCompare: false) == 0)
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
															if (!flag)
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
																presentationShowSubsections = Behavior.GetPresentationShowSubsections(pres);
																flag = true;
															}
															if (presentationShowSubsections)
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
																num++;
																A(shape, shape5, num, 2);
															}
														}
													}
												}
												shape6 = null;
											}
											catch (Exception ex7)
											{
												ProjectData.SetProjectError(ex7);
												Exception ex8 = ex7;
												ProjectData.ClearProjectError();
											}
										}
										while (true)
										{
											switch (3)
											{
											case 0:
												break;
											default:
												goto end_IL_08b4;
											}
											continue;
											end_IL_08b4:
											break;
										}
									}
									finally
									{
										if (enumerator7 is IDisposable)
										{
											while (true)
											{
												switch (3)
												{
												case 0:
													continue;
												}
												(enumerator7 as IDisposable).Dispose();
												break;
											}
										}
									}
								}
								while (true)
								{
									switch (4)
									{
									case 0:
										break;
									default:
										goto end_IL_08ee;
									}
									continue;
									end_IL_08ee:
									break;
								}
							}
							finally
							{
								if (enumerator6 is IDisposable)
								{
									while (true)
									{
										switch (3)
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
					}
				}
				else
				{
					Forms.WarningMessage(AH.A(6010));
				}
				goto IL_092b;
			}
			IL_092b:
			slide = null;
			customLayout = null;
			customLayout2 = null;
			dictionary = null;
			D = null;
			Events.RefreshPresentation(pres);
			return;
			IL_066b:
			B = null;
			C = null;
			goto IL_092b;
		}
	}

	private static List<QB> B(Microsoft.Office.Interop.PowerPoint.Presentation A, CustomLayout B, ref bool C)
	{
		List<QB> list = new List<QB>();
		SectionProperties sectionProperties = A.SectionProperties;
		int count = sectionProperties.Count;
		int num = count;
		checked
		{
			int num7 = default(int);
			for (int num2 = 2; num2 <= num; num2++)
			{
				int num3 = sectionProperties.FirstSlide(num2);
				Slide slide;
				if (num3 != -1)
				{
					if (PowerPointAddIn1.Slides.Helpers.GetSlideType(A.Slides[num3]) == SlideType.Flysheet)
					{
						slide = A.Slides[num3];
						goto IL_0118;
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
				}
				if (num3 == -1)
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
					if (num2 < count)
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
						int num4 = num2;
						int num5 = count;
						int num6 = num4;
						while (true)
						{
							if (num6 <= num5)
							{
								if (sectionProperties.SlidesCount(num6) > 0)
								{
									num7 = sectionProperties.FirstSlide(num6) - 1;
									break;
								}
								num6++;
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
							break;
						}
					}
					else
					{
						num7 = A.Slides.Count;
					}
				}
				else
				{
					num7 = num3 - 1;
				}
				slide = A.Slides.AddSlide(num7 + 1, B);
				Create.SkipPresentationNewSlideEvent(slide);
				slide.MoveToSectionStart(num2);
				goto IL_0118;
				IL_0118:
				QB item = new QB
				{
					A = slide
				};
				if (!Update.A(sectionProperties.Name(num2)))
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
					item.A = 1;
					item.A = sectionProperties.Name(num2);
				}
				else
				{
					item.A = 2;
					item.A = Strings.Mid(sectionProperties.Name(num2), 2);
					C = true;
				}
				list.Add(item);
				slide = null;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				sectionProperties = null;
				return list;
			}
		}
	}

	private static int A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		int result;
		try
		{
			result = Conversions.ToInteger(A.Tags[Constants.TAG_TOC_LEVEL]);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = 0;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static void A(CustomLayout A, Slide B, Microsoft.Office.Interop.PowerPoint.Shape C, string D, int E, int F, int G, ref float H, float I, float J)
	{
		if (F == 0)
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
			J = 0f;
		}
		Microsoft.Office.Interop.PowerPoint.Shape shape = Update.A(B, A, G, D: false, ref H, I, J);
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = shape.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				if (shape2.HasTextFrame == MsoTriState.msoFalse)
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
				TextRange2 textRange = shape2.TextFrame2.TextRange;
				string text = textRange.Text.ToUpper();
				uint num = YG.A(text);
				if (num <= 211792314)
				{
					if (num != 7274958)
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
						if (num != 149826459)
						{
							if (num == 211792314)
							{
								if (Operators.CompareString(text, AH.A(7299), TextCompare: false) == 0)
								{
									goto IL_0251;
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
						else if (Operators.CompareString(text, AH.A(7342), TextCompare: false) == 0)
						{
							goto IL_0251;
						}
					}
					else
					{
						if (Operators.CompareString(text, AH.A(7312), TextCompare: false) == 0)
						{
							goto IL_0251;
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
				else if (num <= 2063731641)
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
					if (num != 2050578930)
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
						if (num == 2063731641)
						{
							if (Operators.CompareString(text, AH.A(7365), TextCompare: false) != 0)
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
								TextRange2 textRange2 = textRange;
								object text2;
								if (E <= 0)
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
									text2 = "";
								}
								else
								{
									text2 = E.ToString();
								}
								textRange2.Text = (string)text2;
							}
						}
					}
					else if (Operators.CompareString(text, AH.A(7277), TextCompare: false) != 0)
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
					else if (D.Length > 0)
					{
						textRange.Text = D;
					}
					else
					{
						shape2.Visible = MsoTriState.msoFalse;
					}
				}
				else if (num != 2381896112u)
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
					if (num != 3278918548u)
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
					else if (Operators.CompareString(text, AH.A(7327), TextCompare: false) == 0)
					{
						goto IL_0251;
					}
				}
				else
				{
					if (Operators.CompareString(text, AH.A(7284), TextCompare: false) == 0)
					{
						goto IL_0251;
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
				goto IL_0292;
				IL_0251:
				textRange.Text = C.TextFrame2.TextRange.Text;
				goto IL_0292;
				IL_0292:
				textRange = null;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_02a1;
				}
				continue;
				end_IL_02a1:
				break;
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
		shape.Visible = MsoTriState.msoTrue;
		shape = null;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Shape B, int C, int D)
	{
		TextRange textRange = A.TextFrame.TextRange.Paragraphs(C, 1);
		textRange.IndentLevel = D;
		textRange.Text = B.TextFrame2.TextRange.Text + AH.A(7894);
		_ = null;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, Slide B)
	{
		int num = 0;
		float E = 0f;
		float f = 0f;
		float g = 0f;
		CustomLayout customLayout = B.CustomLayout;
		try
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape = customLayout.Shapes[Constants.AGENDA_TITLE_LEVEL_1];
			E = shape.Top;
			f = shape.Left;
			_ = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			g = customLayout.Shapes[Constants.AGENDA_TITLE_LEVEL_1_ACTIVE].Top - E - customLayout.Shapes[Constants.AGENDA_TITLE_LEVEL_1].Height;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		customLayout = null;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Slides.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Slide slide = (Slide)enumerator.Current;
				try
				{
					if (slide.Shapes.HasTitle != MsoTriState.msoTrue)
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (PowerPointAddIn1.Slides.Helpers.IsSpecialSlide(slide))
						{
							break;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							Microsoft.Office.Interop.PowerPoint.Shape title = slide.Shapes.Title;
							if (title.HasTextFrame != MsoTriState.msoTrue)
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
								if (title.TextFrame2.HasText != MsoTriState.msoTrue)
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
									num = checked(num + 1);
									Update.A(B, title.TextFrame2.TextRange.Text, Update.A(slide), num, ref E, f, g);
									break;
								}
								break;
							}
							break;
						}
						break;
					}
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
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
	}

	private static void A(Slide A, string B, int C, int D, ref float E, float F, float G)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = null;
		if (D == 0)
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
			G = 0f;
		}
		Microsoft.Office.Interop.PowerPoint.Shape shape3 = Update.A(A, A.CustomLayout, 1, D: false, ref E, F, G);
		shape3.Visible = MsoTriState.msoTrue;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = shape3.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape4 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				if (shape4.HasTextFrame == MsoTriState.msoFalse)
				{
					continue;
				}
				TextRange2 textRange = shape4.TextFrame2.TextRange;
				string text = textRange.Text.ToUpper();
				uint num = YG.A(text);
				if (num <= 211792314)
				{
					if (num != 7274958)
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
						if (num != 149826459)
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
							if (num == 211792314 && Operators.CompareString(text, AH.A(7299), TextCompare: false) == 0)
							{
								goto IL_023f;
							}
						}
						else if (Operators.CompareString(text, AH.A(7342), TextCompare: false) == 0)
						{
							goto IL_023f;
						}
					}
					else if (Operators.CompareString(text, AH.A(7312), TextCompare: false) == 0)
					{
						goto IL_023f;
					}
				}
				else if (num <= 2063731641)
				{
					if (num != 2050578930)
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
						if (num != 2063731641)
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
						else if (Operators.CompareString(text, AH.A(7365), TextCompare: false) != 0)
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
							textRange.Text = ((C > 0) ? C.ToString() : "");
						}
					}
					else if (Operators.CompareString(text, AH.A(7277), TextCompare: false) != 0)
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
						shape4.Visible = MsoTriState.msoFalse;
						shape = shape4;
					}
				}
				else if (num != 2381896112u)
				{
					if (num != 3278918548u)
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
						if (Operators.CompareString(text, AH.A(7327), TextCompare: false) == 0)
						{
							goto IL_023f;
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
					if (Operators.CompareString(text, AH.A(7284), TextCompare: false) == 0)
					{
						goto IL_023f;
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
				goto IL_0267;
				IL_023f:
				textRange.Text = B;
				shape2 = shape4;
				goto IL_0267;
				IL_0267:
				textRange = null;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_0277;
				}
				continue;
				end_IL_0277:
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
		if (shape2 != null)
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
			if (shape != null)
			{
				shape2.Left = shape.Left;
				shape2.Width += shape.Width;
				shape.Delete();
				shape2 = null;
				shape = null;
			}
		}
		shape3 = null;
	}

	private static void A(CustomLayout A, ref float B, ref float C, ref Dictionary<int, float> D)
	{
		CustomLayout customLayout = A;
		try
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape = customLayout.Shapes[Constants.AGENDA_TITLE_LEVEL_1];
			B = shape.Top;
			D[1] = shape.Left;
			shape = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			D[2] = customLayout.Shapes[Constants.AGENDA_TITLE_LEVEL_2].Left;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		try
		{
			D[3] = customLayout.Shapes[Constants.AGENDA_TITLE_LEVEL_3].Left;
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
		try
		{
			C = customLayout.Shapes[Constants.AGENDA_TITLE_LEVEL_1_ACTIVE].Top - B - customLayout.Shapes[Constants.AGENDA_TITLE_LEVEL_1].Height;
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
		customLayout = null;
	}

	private static int A(Slide A)
	{
		int num = 0;
		if (A != null)
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
			Microsoft.Office.Interop.PowerPoint.Shapes shapes = A.Shapes;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = shapes.GetEnumerator();
				while (true)
				{
					if (enumerator.MoveNext())
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
						if (shape.Type == MsoShapeType.msoPlaceholder)
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
							if (shape.Visible == MsoTriState.msoTrue)
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
								if (shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderSlideNumber)
								{
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										if (shape.TextFrame2.HasText == MsoTriState.msoTrue)
										{
											try
											{
												num = Conversions.ToInteger(shape.TextFrame2.TextRange.Text);
											}
											catch (Exception ex)
											{
												ProjectData.SetProjectError(ex);
												Exception ex2 = ex;
												ProjectData.ClearProjectError();
											}
										}
										break;
									}
									break;
								}
							}
						}
						shape = null;
						continue;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_00d5;
						}
						continue;
						end_IL_00d5:
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
						switch (7)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			if (num == 0)
			{
				try
				{
					shapes = ((Microsoft.Office.Interop.PowerPoint.Presentation)A.Parent).Designs[1].SlideMaster.Shapes;
					IEnumerator enumerator2 = default(IEnumerator);
					try
					{
						enumerator2 = shapes.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
							if (shape2.HasTextFrame != MsoTriState.msoTrue || shape2.Type == MsoShapeType.msoPlaceholder || !shape2.TextFrame2.TextRange.Text.Contains(AH.A(7899)))
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
								num = A.SlideNumber;
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
								switch (4)
								{
								case 0:
									continue;
								}
								(enumerator2 as IDisposable).Dispose();
								break;
							}
						}
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
			}
			shapes = null;
		}
		return num;
	}

	internal static bool A(string A)
	{
		return A.StartsWith(Constants.SUBSECTION_PREFIX);
	}

	internal static void A(CustomLayout A, ref List<string> B, ref List<string> C)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Shapes.Placeholders.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				if (shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody)
				{
					TextRange2 textRange = shape.TextFrame2.TextRange;
					if (textRange.get_Paragraphs(-1, -1).Count == 2)
					{
						B = Update.A(textRange.get_Paragraphs(1, -1).ParagraphFormat.Bullet.Style);
						C = Update.A(textRange.get_Paragraphs(2, -1).ParagraphFormat.Bullet.Style);
						return;
					}
					textRange = null;
				}
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
				return;
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

	private static List<string> A(MsoNumberedBulletStyle A)
	{
		List<string> list;
		switch (A)
		{
		case MsoNumberedBulletStyle.msoBulletAlphaUCPeriod:
		case MsoNumberedBulletStyle.msoBulletAlphaUCParenBoth:
		case MsoNumberedBulletStyle.msoBulletAlphaUCParenRight:
			list = new List<string>(new string[50]
			{
				AH.A(7902),
				AH.A(7905),
				AH.A(7908),
				AH.A(7911),
				AH.A(7914),
				AH.A(7917),
				AH.A(7920),
				AH.A(7923),
				AH.A(7926),
				AH.A(7929),
				AH.A(7932),
				AH.A(7935),
				AH.A(7938),
				AH.A(7941),
				AH.A(7944),
				AH.A(7947),
				AH.A(7950),
				AH.A(7953),
				AH.A(7956),
				AH.A(7959),
				AH.A(7962),
				AH.A(7965),
				AH.A(7968),
				AH.A(7971),
				AH.A(7974),
				AH.A(7977),
				AH.A(7980),
				AH.A(7985),
				AH.A(7990),
				AH.A(7995),
				AH.A(8000),
				AH.A(8005),
				AH.A(8010),
				AH.A(8015),
				AH.A(8020),
				AH.A(8025),
				AH.A(8030),
				AH.A(8035),
				AH.A(8040),
				AH.A(8045),
				AH.A(8050),
				AH.A(8055),
				AH.A(8060),
				AH.A(8065),
				AH.A(8070),
				AH.A(8075),
				AH.A(8080),
				AH.A(8085),
				AH.A(8090),
				AH.A(8095)
			});
			break;
		case MsoNumberedBulletStyle.msoBulletAlphaLCPeriod:
		case MsoNumberedBulletStyle.msoBulletAlphaLCParenBoth:
		case MsoNumberedBulletStyle.msoBulletAlphaLCParenRight:
			list = new List<string>(new string[50]
			{
				AH.A(8100),
				AH.A(8103),
				AH.A(8106),
				AH.A(8109),
				AH.A(8112),
				AH.A(8115),
				AH.A(8118),
				AH.A(8121),
				AH.A(8124),
				AH.A(8127),
				AH.A(8130),
				AH.A(8133),
				AH.A(8136),
				AH.A(8139),
				AH.A(8142),
				AH.A(8145),
				AH.A(8148),
				AH.A(8151),
				AH.A(8154),
				AH.A(8157),
				AH.A(8160),
				AH.A(8163),
				AH.A(8166),
				AH.A(8169),
				AH.A(8172),
				AH.A(8175),
				AH.A(8178),
				AH.A(8183),
				AH.A(8188),
				AH.A(8193),
				AH.A(8198),
				AH.A(8203),
				AH.A(8208),
				AH.A(8213),
				AH.A(8218),
				AH.A(8223),
				AH.A(8228),
				AH.A(8233),
				AH.A(8238),
				AH.A(8243),
				AH.A(8248),
				AH.A(8253),
				AH.A(8258),
				AH.A(8263),
				AH.A(8268),
				AH.A(8273),
				AH.A(8278),
				AH.A(8283),
				AH.A(8288),
				AH.A(8293)
			});
			break;
		case MsoNumberedBulletStyle.msoBulletRomanUCPeriod:
		case MsoNumberedBulletStyle.msoBulletRomanUCParenBoth:
		case MsoNumberedBulletStyle.msoBulletRomanUCParenRight:
			list = new List<string>(new string[50]
			{
				AH.A(7926),
				AH.A(8020),
				AH.A(8298),
				AH.A(8305),
				AH.A(7965),
				AH.A(8310),
				AH.A(8315),
				AH.A(8322),
				AH.A(8331),
				AH.A(7971),
				AH.A(8336),
				AH.A(8341),
				AH.A(8348),
				AH.A(8357),
				AH.A(8364),
				AH.A(8369),
				AH.A(8376),
				AH.A(8385),
				AH.A(8396),
				AH.A(8095),
				AH.A(8403),
				AH.A(8410),
				AH.A(8419),
				AH.A(8430),
				AH.A(8439),
				AH.A(8446),
				AH.A(8455),
				AH.A(8466),
				AH.A(8479),
				AH.A(8488),
				AH.A(8495),
				AH.A(8504),
				AH.A(8515),
				AH.A(8528),
				AH.A(8539),
				AH.A(8548),
				AH.A(8559),
				AH.A(8572),
				AH.A(8587),
				AH.A(8598),
				AH.A(8603),
				AH.A(8610),
				AH.A(8619),
				AH.A(8630),
				AH.A(8639),
				AH.A(8646),
				AH.A(8655),
				AH.A(8666),
				AH.A(8679),
				AH.A(7935)
			});
			break;
		case MsoNumberedBulletStyle.msoBulletRomanLCParenBoth:
		case MsoNumberedBulletStyle.msoBulletRomanLCParenRight:
		case MsoNumberedBulletStyle.msoBulletRomanLCPeriod:
			list = new List<string>(new string[50]
			{
				AH.A(8124),
				AH.A(8218),
				AH.A(8688),
				AH.A(8695),
				AH.A(8163),
				AH.A(8700),
				AH.A(8705),
				AH.A(8712),
				AH.A(8721),
				AH.A(8169),
				AH.A(8726),
				AH.A(8731),
				AH.A(8738),
				AH.A(8747),
				AH.A(8754),
				AH.A(8759),
				AH.A(8766),
				AH.A(8775),
				AH.A(8786),
				AH.A(8293),
				AH.A(8793),
				AH.A(8800),
				AH.A(8809),
				AH.A(8820),
				AH.A(8829),
				AH.A(8836),
				AH.A(8845),
				AH.A(8856),
				AH.A(8869),
				AH.A(8878),
				AH.A(8885),
				AH.A(8894),
				AH.A(8905),
				AH.A(8918),
				AH.A(8929),
				AH.A(8938),
				AH.A(8949),
				AH.A(8962),
				AH.A(8977),
				AH.A(8988),
				AH.A(8993),
				AH.A(9000),
				AH.A(9009),
				AH.A(9020),
				AH.A(9029),
				AH.A(9036),
				AH.A(9045),
				AH.A(9056),
				AH.A(9069),
				AH.A(8133)
			});
			break;
		default:
			list = new List<string>(new string[50]
			{
				AH.A(9078),
				AH.A(9081),
				AH.A(9084),
				AH.A(9087),
				AH.A(9090),
				AH.A(9093),
				AH.A(9096),
				AH.A(9099),
				AH.A(9102),
				AH.A(9105),
				AH.A(9110),
				AH.A(9115),
				AH.A(9120),
				AH.A(9125),
				AH.A(9130),
				AH.A(9135),
				AH.A(9140),
				AH.A(9145),
				AH.A(9150),
				AH.A(9155),
				AH.A(9160),
				AH.A(9165),
				AH.A(9170),
				AH.A(9175),
				AH.A(9180),
				AH.A(9185),
				AH.A(9190),
				AH.A(9195),
				AH.A(9200),
				AH.A(9205),
				AH.A(9210),
				AH.A(9215),
				AH.A(9220),
				AH.A(9225),
				AH.A(9230),
				AH.A(9235),
				AH.A(9240),
				AH.A(9245),
				AH.A(9250),
				AH.A(9255),
				AH.A(9260),
				AH.A(9265),
				AH.A(9270),
				AH.A(9275),
				AH.A(9280),
				AH.A(9285),
				AH.A(9290),
				AH.A(9295),
				AH.A(9300),
				AH.A(9305)
			});
			break;
		}
		if (KG.A.ShowBulletPunctuation)
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
			switch (A)
			{
			case MsoNumberedBulletStyle.msoBulletAlphaLCPeriod:
			case MsoNumberedBulletStyle.msoBulletAlphaUCPeriod:
			case MsoNumberedBulletStyle.msoBulletArabicPeriod:
			case MsoNumberedBulletStyle.msoBulletRomanLCPeriod:
			case MsoNumberedBulletStyle.msoBulletRomanUCPeriod:
			{
				List<string> source2 = list;
				Func<string, string> selector2;
				if (_Closure_0024__.A == null)
				{
					selector2 = (_Closure_0024__.A = [SpecialName] (string text) => text + AH.A(14417));
				}
				else
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
					selector2 = _Closure_0024__.A;
				}
				list = source2.Select(selector2).ToList();
				break;
			}
			case MsoNumberedBulletStyle.msoBulletArabicParenRight:
			case MsoNumberedBulletStyle.msoBulletRomanLCParenRight:
			case MsoNumberedBulletStyle.msoBulletAlphaLCParenRight:
			case MsoNumberedBulletStyle.msoBulletAlphaUCParenRight:
			case MsoNumberedBulletStyle.msoBulletRomanUCParenRight:
			{
				List<string> source = list;
				Func<string, string> selector;
				if (_Closure_0024__.B == null)
				{
					selector = (_Closure_0024__.B = [SpecialName] (string text) => text + AH.A(14255));
				}
				else
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
					selector = _Closure_0024__.B;
				}
				list = source.Select(selector).ToList();
				break;
			}
			case MsoNumberedBulletStyle.msoBulletRomanLCParenBoth:
			case MsoNumberedBulletStyle.msoBulletAlphaLCParenBoth:
			case MsoNumberedBulletStyle.msoBulletAlphaUCParenBoth:
			case MsoNumberedBulletStyle.msoBulletArabicParenBoth:
			case MsoNumberedBulletStyle.msoBulletRomanUCParenBoth:
				list = list.Select([SpecialName] (string text) => AH.A(17795) + text + AH.A(14255)).ToList();
				break;
			}
		}
		return list;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, List<QB> B, SlideType C, Slide D)
	{
		RB a = default(RB);
		RB CS_0024_003C_003E8__locals9 = new RB(a);
		int num = 0;
		if (D != null)
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
			num = D.SlideIndex;
		}
		for (int i = A.Slides.Count; i >= 1; i = checked(i + -1))
		{
			CS_0024_003C_003E8__locals9.A = A.Slides[i];
			if (PowerPointAddIn1.Slides.Helpers.GetSlideType(CS_0024_003C_003E8__locals9.A) != C || CS_0024_003C_003E8__locals9.A.SlideIndex == num)
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
			Func<QB, bool> predicate;
			if (CS_0024_003C_003E8__locals9.A != null)
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
				predicate = CS_0024_003C_003E8__locals9.A;
			}
			else
			{
				predicate = (CS_0024_003C_003E8__locals9.A = [SpecialName] (QB qB) =>
				{
					if (qB.A != null)
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
								return qB.A.SlideIndex == CS_0024_003C_003E8__locals9.A.SlideIndex;
							}
						}
					}
					return false;
				});
			}
			if (B.Where(predicate).ToList().Any())
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
			CS_0024_003C_003E8__locals9.A.Delete();
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			CS_0024_003C_003E8__locals9.A = null;
			return;
		}
	}

	private static void A(Slide A)
	{
		for (int i = A.Shapes.Count; i >= 1; i = checked(i + -1))
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape = A.Shapes[i];
			if (shape.Type == MsoShapeType.msoGroup)
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
				if (shape.Name.ToLower().StartsWith(AH.A(9310)))
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
					shape.Delete();
				}
			}
			shape = null;
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(Slide A, CustomLayout B, int C, bool D, ref float E, float F, float G)
	{
		string text = AH.A(9323);
		string text2;
		if (!D)
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
			text2 = C.ToString();
		}
		else
		{
			text2 = C + AH.A(9350);
		}
		string index = text + text2;
		B.Shapes[index].Copy();
		Microsoft.Office.Interop.PowerPoint.Shape shape = default(Microsoft.Office.Interop.PowerPoint.Shape);
		try
		{
			shape = A.Shapes.Paste()[1];
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			int num = 0;
			bool flag = false;
			while (num < 10)
			{
				try
				{
					B.Shapes[index].Copy();
					Thread.Sleep(10);
					System.Windows.Forms.Application.DoEvents();
					shape = A.Shapes.Paste()[1];
					flag = true;
					clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)11, AH.A(9365));
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					num = checked(num + 1);
					ProjectData.ClearProjectError();
					continue;
				}
				break;
			}
			if (!flag)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					if (ex2.Message.Contains(AH.A(9428)))
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
						Forms.ErrorMessage(AH.A(9465));
					}
					clsReporting.LogException(ex2);
					throw;
				}
			}
			ProjectData.ClearProjectError();
		}
		shape.Top = E;
		shape.Left = F;
		E += shape.Height + G;
		return shape;
	}
}
