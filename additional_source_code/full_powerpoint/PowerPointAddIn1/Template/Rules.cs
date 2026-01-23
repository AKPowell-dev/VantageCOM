using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.Libraries;
using MacabacusMacros.Libraries.Caching;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Library2;
using PowerPointAddIn1.Presentation;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Template;

public sealed class Rules
{
	[CompilerGenerated]
	private static Options m_A;

	[CompilerGenerated]
	private static List<LegalNotice> m_A;

	[CompilerGenerated]
	private static Dictionary<Microsoft.Office.Interop.PowerPoint.Presentation, Microsoft.Office.Interop.PowerPoint.Shape> m_A;

	private static Options CheckOptions
	{
		[CompilerGenerated]
		get
		{
			return Rules.m_A;
		}
		[CompilerGenerated]
		set
		{
			Rules.m_A = value;
		}
	} = null;

	private static List<LegalNotice> LegalNotices
	{
		[CompilerGenerated]
		get
		{
			return Rules.m_A;
		}
		[CompilerGenerated]
		set
		{
			Rules.m_A = value;
		}
	} = null;

	private static Dictionary<Microsoft.Office.Interop.PowerPoint.Presentation, Microsoft.Office.Interop.PowerPoint.Shape> LegalShapes
	{
		[CompilerGenerated]
		get
		{
			return Rules.m_A;
		}
		[CompilerGenerated]
		set
		{
			Rules.m_A = value;
		}
	} = null;

	public static void LoadOptions(XmlDocument xmlDoc)
	{
		CheckOptions = new Options(xmlDoc);
		LegalShapes = new Dictionary<Microsoft.Office.Interop.PowerPoint.Presentation, Microsoft.Office.Interop.PowerPoint.Shape>();
		Options checkOptions = CheckOptions;
		if (!checkOptions.RequireLegalSlide)
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
			if (!checkOptions.RequireContactSlide)
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
				if (!checkOptions.RequireFrontCover)
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
					if (!checkOptions.RequireBackCover)
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
						if (checkOptions.LegalCheckAction == Options.LegalCheckActionEnum.DoNotCheck)
						{
							goto IL_0081;
						}
					}
				}
			}
		}
		A();
		goto IL_0081;
		IL_0081:
		checkOptions = null;
	}

	private static void A()
	{
		StopEnforcement();
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(137502)).AddEventHandler(NG.A.Application, new EApplication_PresentationBeforeSaveEventHandler(A));
	}

	public static void StopEnforcement()
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(137502)).RemoveEventHandler(NG.A.Application, new EApplication_PresentationBeforeSaveEventHandler(A));
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, ref bool B)
	{
		using (List<LibraryItem>.Enumerator enumerator = Base.LibraryCollection.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				LibraryItem current = enumerator.Current;
				if (!A.FullName.StartsWith(current.Location))
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return;
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_004f;
				}
				continue;
				end_IL_004f:
				break;
			}
		}
		CustomLayout B2 = null;
		string text = "";
		if (!Rules.A(A, ref B2))
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
			text = AH.A(137547);
		}
		else if (!Rules.A(A))
		{
			text = AH.A(137630);
		}
		else if (!Rules.A(A, B2))
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
			text = AH.A(137832);
		}
		if (text.Length > 0)
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
			if (CheckOptions.RuleViolationAction == Options.RuleViolationEnum.WarnOnSave)
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
				Forms.WarningMessage(text);
			}
			else
			{
				Forms.ErrorMessage(AH.A(137974) + text);
				B = true;
			}
		}
		B2 = null;
	}

	public static bool ProofingCheckRequiredSlides(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		CustomLayout B = null;
		return A(pres, ref B);
	}

	public static bool ProofingCheckCoverPosition(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		return A(pres);
	}

	public static bool ProofingCheckLegalNotices(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		return A(pres, null);
	}

	private static bool A(Microsoft.Office.Interop.PowerPoint.Presentation A, ref CustomLayout B)
	{
		bool result = true;
		bool flag = false;
		bool flag2 = false;
		bool flag3 = false;
		bool flag4 = false;
		bool flag5 = false;
		bool requireLegalSlide = CheckOptions.RequireLegalSlide;
		bool requireContactSlide = CheckOptions.RequireContactSlide;
		bool requireFrontCover = CheckOptions.RequireFrontCover;
		bool requireBackCover = CheckOptions.RequireBackCover;
		if (!requireLegalSlide)
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
			if (!requireContactSlide)
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
				if (!requireFrontCover)
				{
					if (!requireBackCover)
					{
						goto IL_035c;
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
		checked
		{
			try
			{
				int count = A.Slides.Count;
				if (count > 1)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						int num = count;
						int num2 = 1;
						while (true)
						{
							if (num2 <= num)
							{
								if (PowerPointAddIn1.Slides.Helpers.GetSlideType(A.Slides[num2]) == SlideType.Title)
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
									flag = true;
									break;
								}
								if (num2 == 2)
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
										goto end_IL_00de;
									}
									continue;
									end_IL_00de:
									break;
								}
								num2++;
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
							break;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							int num3 = count;
							while (true)
							{
								if (num3 >= 1)
								{
									Slide slide = A.Slides[num3];
									if (requireLegalSlide)
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
										if (!flag2)
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
											if (PowerPointAddIn1.Slides.Helpers.GetSlideType(slide) == SlideType.Legal)
											{
												flag2 = true;
												B = slide.CustomLayout;
											}
										}
									}
									if (requireContactSlide && !flag3)
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
										if (PowerPointAddIn1.Slides.Helpers.GetSlideType(slide) == SlideType.Contact)
										{
											flag3 = true;
										}
									}
									if (requireFrontCover)
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
										if (!flag4 && PowerPointAddIn1.Slides.Helpers.GetSlideType(slide) == SlideType.CoverFront)
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
											flag4 = true;
										}
									}
									if (requireBackCover)
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
										if (!flag5)
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
											if (PowerPointAddIn1.Slides.Helpers.GetSlideType(slide) == SlideType.CoverBack)
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
												flag5 = true;
											}
										}
									}
									if (flag2 && flag3)
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
										if (flag4)
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
											if (flag5)
											{
												break;
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
									if (flag2)
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
										if (!requireContactSlide)
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
											if (!requireFrontCover)
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
												if (!requireBackCover)
												{
													break;
												}
											}
										}
									}
									if (flag3)
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
										if (!requireLegalSlide)
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
											if (!requireFrontCover)
											{
												if (!requireBackCover)
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
											}
										}
									}
									if (flag4 && !requireLegalSlide)
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
										if (!requireContactSlide)
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
											if (!requireBackCover)
											{
												break;
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
									if (flag5 && !requireLegalSlide)
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
										if (!requireContactSlide)
										{
											if (!requireFrontCover)
											{
												break;
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
									slide = null;
									num3 += -1;
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
							if (requireLegalSlide)
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
								if (!flag2)
								{
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										result = false;
										break;
									}
									break;
								}
							}
							if (requireContactSlide)
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
								if (!flag3)
								{
									result = false;
									break;
								}
							}
							if (requireFrontCover)
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
								if (!flag4)
								{
									while (true)
									{
										switch (5)
										{
										case 0:
											continue;
										}
										result = false;
										break;
									}
									break;
								}
							}
							if (!requireBackCover || flag5)
							{
								break;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								result = false;
								break;
							}
							break;
						}
						break;
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			goto IL_035c;
		}
		IL_035c:
		return result;
	}

	private static bool A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		bool result = true;
		try
		{
			if (A.Slides.Count > 0)
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
					if (CheckOptions.RequireFrontCover)
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
						if (PowerPointAddIn1.Slides.Helpers.GetSlideType(A.Slides[1]) != SlideType.CoverFront)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								result = false;
								break;
							}
							break;
						}
					}
					if (CheckOptions.RequireBackCover && PowerPointAddIn1.Slides.Helpers.GetSlideType(A.Slides[A.Slides.Count]) != SlideType.CoverBack)
					{
						result = false;
					}
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static bool A(Microsoft.Office.Interop.PowerPoint.Presentation A, CustomLayout B)
	{
		bool flag = true;
		bool flag2 = false;
		Microsoft.Office.Interop.PowerPoint.Shape value = null;
		checked
		{
			if (CheckOptions.LegalCheckAction != Options.LegalCheckActionEnum.DoNotCheck)
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
				try
				{
					if (LegalShapes.TryGetValue(A, out value))
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
						try
						{
							_ = value.Name;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							LegalShapes.Remove(A);
							value = null;
							ProjectData.ClearProjectError();
						}
					}
					if (value == null)
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
						flag2 = true;
						if (B == null)
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
							int count = A.Slides.Count;
							int num = count;
							while (true)
							{
								if (num >= 1)
								{
									if (PowerPointAddIn1.Slides.Helpers.GetSlideType(A.Slides[num]) == SlideType.Legal)
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
										B = A.Slides[num].CustomLayout;
									}
									if (num == count - 2)
									{
										break;
									}
									while (true)
									{
										switch (6)
										{
										case 0:
											break;
										default:
											goto end_IL_00f0;
										}
										continue;
										end_IL_00f0:
										break;
									}
									num += -1;
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
						if (B != null)
						{
							IEnumerator enumerator = default(IEnumerator);
							try
							{
								enumerator = B.Shapes.GetEnumerator();
								while (true)
								{
									if (enumerator.MoveNext())
									{
										Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
										if (Rules.A(shape, A))
										{
											value = shape;
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
											goto end_IL_0149;
										}
										continue;
										end_IL_0149:
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
							B = null;
						}
					}
					if (value != null)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							if (flag2)
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
								LegalShapes.Add(A, value);
							}
							if (LegalNotices == null)
							{
								Options.LegalCheckActionEnum legalCheckAction = CheckOptions.LegalCheckAction;
								if (legalCheckAction != Options.LegalCheckActionEnum.CheckAgainstOriginalTemplate)
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
									if (legalCheckAction != Options.LegalCheckActionEnum.CheckAgainstAllTemplates)
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
										LegalNotices = Rules.A();
									}
								}
								else
								{
									LegalNotices = Rules.A(A);
								}
							}
							if (LegalNotices == null)
							{
								break;
							}
							using (List<LegalNotice>.Enumerator enumerator2 = LegalNotices.GetEnumerator())
							{
								while (true)
								{
									if (enumerator2.MoveNext())
									{
										LegalNotice current = enumerator2.Current;
										TextRange2 textRange = value.TextFrame2.TextRange;
										if (textRange.get_Paragraphs(-1, -1).Count != current.Paragraphs)
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
											flag = false;
										}
										else if (textRange.get_Lines(-1, -1).Count != current.Lines)
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
											flag = false;
										}
										else if (textRange.Text.Length != current.Length)
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
											flag = false;
										}
										else if (Operators.CompareString(textRange.Text, current.Text, TextCompare: false) != 0)
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
											flag = false;
										}
										textRange = null;
										if (flag)
										{
											break;
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
										continue;
									}
									while (true)
									{
										switch (2)
										{
										case 0:
											break;
										default:
											goto end_IL_02f0;
										}
										continue;
										end_IL_02f0:
										break;
									}
									break;
								}
							}
							break;
						}
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					clsReporting.LogException(ex4);
					ProjectData.ClearProjectError();
				}
			}
			return flag;
		}
	}

	private static List<LegalNotice> A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		List<LegalNotice> list = new List<LegalNotice>();
		Microsoft.Office.Interop.PowerPoint.Presentation A2 = null;
		bool B = false;
		string templateId = Templates.GetTemplateId(A);
		if (templateId.Length > 0)
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
			Templates.B(ref A2, ref B, templateId, A.Application);
			if (A2 != null)
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
				if (Rules.A(A2) != null)
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
					list.Add(Rules.A(A2));
				}
				if (B)
				{
					PowerPointAddIn1.Presentation.Helpers.CloseQuietly(A2);
				}
				A2 = null;
			}
		}
		return list;
	}

	private static List<LegalNotice> A()
	{
		//IL_006b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0070: Unknown result type (might be due to invalid IL or missing references)
		//IL_0073: Invalid comparison between Unknown and I4
		//IL_009a: Unknown result type (might be due to invalid IL or missing references)
		//IL_009f: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a2: Invalid comparison between Unknown and I4
		List<LegalNotice> list = new List<LegalNotice>();
		bool flag = Base.IsUserAdmin();
		using List<LibraryItem>.Enumerator enumerator = Base.LibraryCollection.GetEnumerator();
		while (enumerator.MoveNext())
		{
			LibraryItem current = enumerator.Current;
			if (!Access.UserHasAccess(current, (AccessType)0, flag) || !current.IsSharedAndInternallyManaged())
			{
				continue;
			}
			if (CheckOptions.LegalCheckScope != Options.LegalCheckScopeEnum.Both)
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
				if ((int)current.Permission == 2)
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
					if (CheckOptions.LegalCheckScope == Options.LegalCheckScopeEnum.Restricted)
					{
						goto IL_00bf;
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
				if ((int)current.Permission != 1)
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
				if (CheckOptions.LegalCheckScope != Options.LegalCheckScopeEnum.Unrestricted)
				{
					continue;
				}
			}
			goto IL_00bf;
			IL_00bf:
			list.AddRange(A(current));
		}
		while (true)
		{
			switch (3)
			{
			case 0:
				continue;
			}
			return list;
		}
	}

	private static List<LegalNotice> A(LibraryItem A)
	{
		List<LegalNotice> list = new List<LegalNotice>();
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
		string text = A.Location;
		if (clsCaching.UseCachedLibraries(text))
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
			text = clsCaching.CachedLibraryFolder(A);
		}
		if (Directory.Exists(text))
		{
			string path = Path.Combine(text, Base.LIB_MASTERS_FOLDER_NAME);
			if (Directory.Exists(path))
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
				string[] directories = Directory.GetDirectories(path);
				foreach (string path2 in directories)
				{
					try
					{
						FileInfo[] files = new DirectoryInfo(path2).GetFiles(AH.A(138005), SearchOption.TopDirectoryOnly);
						foreach (FileInfo fileInfo in files)
						{
							try
							{
								presentation = PowerPointAddIn1.Presentation.Helpers.OpenQuietly(NG.A.Application, fileInfo.FullName);
								if (Rules.A(presentation) != null)
								{
									list.Add(Rules.A(presentation));
								}
								PowerPointAddIn1.Presentation.Helpers.CloseQuietly(presentation);
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
							finally
							{
								presentation = null;
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0108;
							}
							continue;
							end_IL_0108:
							break;
						}
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
					switch (4)
					{
					case 0:
						continue;
					}
					break;
				}
			}
		}
		return list;
	}

	private static LegalNotice A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Designs.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				CustomLayouts customLayouts = ((Design)enumerator.Current).SlideMaster.CustomLayouts;
				int num = customLayouts.Count;
				while (true)
				{
					if (num >= 1)
					{
						if (PowerPointAddIn1.Slides.Helpers.GetLayoutType(customLayouts[num]) == SlideType.Legal)
						{
							try
							{
								enumerator2 = customLayouts[num].Shapes.GetEnumerator();
								while (enumerator2.MoveNext())
								{
									Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
									if (!Rules.A(shape, A))
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
										if (1 == 0)
										{
											/*OpCode not supported: LdMemberToken*/;
										}
										return new LegalNotice(shape);
									}
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_00b3;
									}
									continue;
									end_IL_00b3:
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
							break;
						}
						num = checked(num + -1);
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
					break;
				}
				customLayouts = null;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_0103;
				}
				continue;
				end_IL_0103:
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
		return null;
	}

	private static bool A(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Presentation B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		if (shape.HasTextFrame == MsoTriState.msoTrue)
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
			if (shape.TextFrame2.HasText == MsoTriState.msoTrue)
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
				if (shape.TextFrame2.TextRange.get_Lines(-1, -1).Count >= 5)
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
					if (shape.TextFrame2.TextRange.Text.Length >= 300)
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
						if ((double)shape.Width >= 0.75 * (double)B.PageSetup.SlideWidth)
						{
							return true;
						}
					}
				}
			}
		}
		shape = null;
		return false;
	}
}
