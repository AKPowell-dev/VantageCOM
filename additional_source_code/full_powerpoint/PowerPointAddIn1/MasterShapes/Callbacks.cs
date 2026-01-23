using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.MasterShapes;

public sealed class Callbacks
{
	public static string Menu()
	{
		Application application = NG.A.Application;
		StringBuilder stringBuilder = new StringBuilder(AH.A(47526));
		bool flag = false;
		List<string> list;
		if (application.Windows.Count > 0)
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
			Base.A(application);
			Stamps.ConvertLegacyStamps();
			SectionTitles.ConvertLegacySectionTitles();
			Base.C();
			list = new List<string>();
			list.Add(AH.A(7959));
			list.Add(AH.A(7962));
			list.Add(AH.A(7914));
			stringBuilder.Append(AH.A(140536));
			Behavior behavior = Behavior.SelectedSlides;
			using (Dictionary<string, Microsoft.Office.Interop.PowerPoint.Shape>.Enumerator enumerator = Base.MyMasterShapes.GetEnumerator())
			{
				string text2 = default(string);
				while (enumerator.MoveNext())
				{
					KeyValuePair<string, Microsoft.Office.Interop.PowerPoint.Shape> current = enumerator.Current;
					string text = Regex.Replace(current.Value.Name, Base.REGEX_PATTERN, "").Trim();
					Behavior behavior2 = Base.A(current.Value.Name);
					switch (behavior2)
					{
					case Behavior.SelectedSlides:
						text2 = AH.A(141464);
						break;
					case Behavior.AllSlides:
						text2 = AH.A(141563);
						break;
					case Behavior.AllLayouts:
						text2 = AH.A(141682);
						break;
					case Behavior.ContentSlides:
						text2 = AH.A(142046);
						break;
					case Behavior.ContentLayouts:
						text2 = AH.A(142173);
						break;
					case Behavior.SlidesShowingBackgroundGraphics:
						text2 = AH.A(142545);
						break;
					case Behavior.LayoutsShowingBackgroundGraphics:
						text2 = AH.A(142735);
						break;
					case Behavior.DynamicSlides:
						text2 = AH.A(143169);
						break;
					case Behavior.DynamicLayouts:
						text2 = AH.A(143445);
						break;
					case Behavior.SpecialSlides:
						text2 = AH.A(143965);
						break;
					case Behavior.SpecialLayouts:
						text2 = AH.A(144349);
						break;
					case Behavior.AboveTopLeft:
						text2 = AH.A(144979);
						break;
					case Behavior.AboveTopRight:
						text2 = AH.A(145159);
						break;
					case Behavior.BelowBottomRight:
						text2 = AH.A(145341);
						break;
					case Behavior.BelowBottomLeft:
						text2 = AH.A(145529);
						break;
					case Behavior.InsideTopLeft:
						text2 = AH.A(145715);
						break;
					case Behavior.InsideTopRight:
						text2 = AH.A(145897);
						break;
					case Behavior.InsideBottomRight:
						text2 = AH.A(146081);
						break;
					case Behavior.InsideBottomLeft:
						text2 = AH.A(146271);
						break;
					case Behavior.CenterInShape:
						text2 = AH.A(146459);
						break;
					}
					if (behavior == Behavior.SelectedSlides)
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
						if (behavior2 != behavior)
						{
							stringBuilder.Append(AH.A(146635));
							flag = true;
						}
					}
					behavior = behavior2;
					string text3 = clsRibbon.GenerateLabel(text, list);
					text = clsRibbon.FixAmpersand(text);
					string text4 = "";
					if (Stamps.HasStampPlaceholder(current.Value))
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
						stringBuilder.Append(AH.A(146714) + current.Key + AH.A(15132) + text4 + AH.A(146771) + text3 + AH.A(146788) + current.Key + AH.A(146851) + text + AH.A(146892) + text2 + AH.A(82681));
					}
					else if (SectionTitles.HasSectionTitlePlaceholder(current.Value))
					{
						stringBuilder.Append(AH.A(146917) + current.Key + AH.A(15132) + text4 + AH.A(146771) + text3 + AH.A(146976) + current.Key + AH.A(146851) + text + AH.A(146892) + text2 + AH.A(82681));
					}
					else if (A(behavior2))
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
						stringBuilder.Append(AH.A(147126) + current.Key + AH.A(15132) + text4 + AH.A(146771) + text3 + AH.A(147173) + current.Key + AH.A(146851) + text + AH.A(146892) + text2 + AH.A(82681));
					}
					else
					{
						stringBuilder.Append(AH.A(146917) + current.Key + AH.A(15132) + text4 + AH.A(146771) + text3 + AH.A(147246) + current.Key + AH.A(146851) + text + AH.A(146892) + text2 + AH.A(82681));
					}
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_0605;
					}
					continue;
					end_IL_0605:
					break;
				}
			}
			if (flag)
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
				stringBuilder.Append(AH.A(147382));
				stringBuilder.Append(AH.A(147471));
			}
			if (Base.IsUserAdmin())
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
				stringBuilder.Append(AH.A(148157));
				stringBuilder.Append(AH.A(148278));
			}
		}
		application = null;
		list = null;
		stringBuilder.Append(AH.A(49007));
		return stringBuilder.ToString();
	}

	private static bool A(Behavior A)
	{
		if ((uint)(A - 11) <= 8u)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return true;
				}
			}
		}
		return false;
	}

	public static bool IsPresent(IRibbonControl control)
	{
		bool flag = false;
		Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Shape value = null;
		string tag = control.Tag;
		if (application.ActivePresentation.Slides.Count > 0)
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
			if (Base.MyMasterShapes.TryGetValue(tag, out value))
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
				Behavior behavior = Base.A(value.Name);
				List<Slide> list = Helpers.A(application);
				if (list.Count > 0 && !Base.A(application, B: false))
				{
					switch (behavior)
					{
					case Behavior.SelectedSlides:
					case Behavior.AllSlides:
					case Behavior.ContentSlides:
					case Behavior.SlidesShowingBackgroundGraphics:
					case Behavior.DynamicSlides:
					case Behavior.SpecialSlides:
					{
						using (List<Slide>.Enumerator enumerator3 = list.GetEnumerator())
						{
							IEnumerator enumerator4 = default(IEnumerator);
							do
							{
								if (enumerator3.MoveNext())
								{
									Slide current2 = enumerator3.Current;
									try
									{
										enumerator4 = current2.Shapes.GetEnumerator();
										while (true)
										{
											if (enumerator4.MoveNext())
											{
												if (Base.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator4.Current, tag))
												{
													flag = true;
													break;
												}
												continue;
											}
											while (true)
											{
												switch (4)
												{
												case 0:
													break;
												default:
													goto end_IL_011e;
												}
												continue;
												end_IL_011e:
												break;
											}
											break;
										}
									}
									finally
									{
										if (enumerator4 is IDisposable)
										{
											while (true)
											{
												switch (1)
												{
												case 0:
													continue;
												}
												(enumerator4 as IDisposable).Dispose();
												break;
											}
										}
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
										goto end_IL_015e;
									}
									continue;
									end_IL_015e:
									break;
								}
								break;
							}
							while (!flag);
						}
						break;
					}
					default:
					{
						using (List<Slide>.Enumerator enumerator = list.GetEnumerator())
						{
							IEnumerator enumerator2 = default(IEnumerator);
							while (true)
							{
								if (enumerator.MoveNext())
								{
									Slide current = enumerator.Current;
									{
										enumerator2 = current.CustomLayout.Shapes.GetEnumerator();
										try
										{
											while (true)
											{
												if (enumerator2.MoveNext())
												{
													if (Base.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, tag))
													{
														flag = true;
														break;
													}
													continue;
												}
												while (true)
												{
													switch (4)
													{
													case 0:
														break;
													default:
														goto end_IL_01d1;
													}
													continue;
													end_IL_01d1:
													break;
												}
												break;
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
									if (!flag)
									{
										continue;
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											break;
										default:
											goto end_IL_01f6;
										}
										continue;
										end_IL_01f6:
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
										goto end_IL_020e;
									}
									continue;
									end_IL_020e:
									break;
								}
								break;
							}
						}
						break;
					}
					}
				}
				list.Clear();
				list = null;
				value = null;
			}
		}
		application = null;
		return flag;
	}
}
