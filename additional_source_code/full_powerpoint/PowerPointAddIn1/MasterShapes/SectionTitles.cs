using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Agenda;
using PowerPointAddIn1.Shapes;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.MasterShapes;

public sealed class SectionTitles
{
	[CompilerGenerated]
	internal sealed class VF
	{
		public Microsoft.Office.Interop.PowerPoint.Shape A;

		public VF(VF A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal void A()
		{
			this.A.Copy();
		}
	}

	[CompilerGenerated]
	internal sealed class WF
	{
		public Microsoft.Office.Interop.PowerPoint.Shape A;

		[SpecialName]
		internal void A()
		{
			this.A.Copy();
		}
	}

	private static readonly string m_A = AH.A(149679);

	private static readonly string m_B = AH.A(149720);

	private static string C = string.Empty;

	private static string A
	{
		get
		{
			return C;
		}
		set
		{
			C = value;
		}
	}

	public static void Toggle(IRibbonControl control, bool blnAdd)
	{
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = NG.A.Application.ActivePresentation;
		if (A(activePresentation))
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
			string strId;
			if (Operators.CompareString(control.Id, AH.A(149225), TextCompare: false) == 0)
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
				Base.C();
				strId = SectionTitles.A;
			}
			else
			{
				strId = control.Tag;
			}
			if (blnAdd)
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
				activePresentation.Tags.Add(SectionTitles.m_B, AH.A(9078));
			}
			else
			{
				activePresentation.Tags.Delete(SectionTitles.m_B);
			}
			AddRemove.Toggle(strId, blnAdd);
		}
		else if (blnAdd)
		{
			A(A: false);
		}
		else
		{
			A(activePresentation);
		}
		if (Operators.CompareString(control.Id, AH.A(149225), TextCompare: false) != 0)
		{
			KG.A.InvalidateControl(AH.A(149225));
		}
		activePresentation = null;
	}

	public static void Remove(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		if (A(pres))
		{
			AddRemove.Toggle(SectionTitles.A, blnAdd: false);
			pres.Tags.Delete(SectionTitles.m_B);
		}
		else
		{
			A(pres);
		}
	}

	public static void Update(Microsoft.Office.Interop.PowerPoint.Presentation pres, bool SuppressMessages, Slide sldNew = null)
	{
		if (!IsVisible(pres))
		{
			return;
		}
		VF a = default(VF);
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
			if (A(pres))
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
					{
						if (sldNew == null)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									AddRemove.Toggle(SectionTitles.A, blnAdd: true);
									return;
								}
							}
						}
						VF CS_0024_003C_003E8__locals8 = new VF(a);
						CS_0024_003C_003E8__locals8.A = null;
						bool flag = false;
						try
						{
							enumerator = pres.Designs[1].SlideMaster.Shapes.GetEnumerator();
							while (true)
							{
								if (!enumerator.MoveNext())
								{
									while (true)
									{
										switch (4)
										{
										case 0:
											break;
										default:
											goto end_IL_00d8;
										}
										continue;
										end_IL_00d8:
										break;
									}
									break;
								}
								Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
								if (Operators.CompareString(shape.Id.ToString(), SectionTitles.A, TextCompare: false) == 0)
								{
									while (true)
									{
										switch (2)
										{
										case 0:
											break;
										default:
											CS_0024_003C_003E8__locals8.A = shape;
											goto end_IL_00bb;
										}
										continue;
										end_IL_00bb:
										break;
									}
									break;
								}
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
										break;
									default:
										(enumerator as IDisposable).Dispose();
										goto end_IL_00ec;
									}
									continue;
									end_IL_00ec:
									break;
								}
							}
						}
						if (CS_0024_003C_003E8__locals8.A != null)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									switch (Base.A(CS_0024_003C_003E8__locals8.A.Name))
									{
									case Behavior.ContentSlides:
										flag = !PowerPointAddIn1.Slides.Helpers.IsSpecialSlide(sldNew);
										break;
									case Behavior.SlidesShowingBackgroundGraphics:
										flag = sldNew.CustomLayout.DisplayMasterShapes == MsoTriState.msoTrue;
										break;
									case Behavior.DynamicSlides:
									{
										SlideType slideType = PowerPointAddIn1.Slides.Helpers.GetSlideType(sldNew);
										if ((uint)(slideType - 4) > 1u)
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
											if ((uint)(slideType - 9) > 1u)
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
												flag = true;
											}
										}
										break;
									}
									case Behavior.AllSlides:
										flag = true;
										break;
									}
									if (flag)
									{
										while (true)
										{
											switch (2)
											{
											case 0:
												break;
											default:
												clsClipboard.CopyWithWait((Action)([SpecialName] () =>
												{
													CS_0024_003C_003E8__locals8.A.Copy();
												}), 4000);
												AddRemove.AddRemove(sldNew, CS_0024_003C_003E8__locals8.A, blnAdd: true, AddRemove.MasterShapeProperties(CS_0024_003C_003E8__locals8.A), pres, DateTime.Now, "");
												CS_0024_003C_003E8__locals8.A.Copy();
												return;
											}
										}
									}
									return;
								}
							}
						}
						return;
					}
					}
				}
			}
			A(SuppressMessages, sldNew);
			return;
		}
	}

	public static bool IsVisible(Microsoft.Office.Interop.PowerPoint.Presentation pres = null)
	{
		if (pres == null)
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
			pres = NG.A.Application.ActivePresentation;
		}
		bool result;
		try
		{
			result = Conversions.ToBoolean(pres.Tags[SectionTitles.m_B]);
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

	private static bool A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Designs[1].SlideMaster.Shapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				if (shape.Visible != MsoTriState.msoFalse)
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
				if (!HasSectionTitlePlaceholder(shape))
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
					SectionTitles.A = shape.Id.ToString();
					return true;
				}
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_008f;
				}
				continue;
				end_IL_008f:
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
		return false;
	}

	public static bool HasSectionTitlePlaceholder(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		bool flag = false;
		if (shp.HasTextFrame == MsoTriState.msoTrue)
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
			Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = shp.TextFrame2;
			int num;
			if (textFrame.HasText == MsoTriState.msoTrue)
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
				if (!textFrame.TextRange.Text.Contains(Placeholders.PLACEHOLDER_SECTION))
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
					num = (textFrame.TextRange.Text.Contains(Placeholders.PLACEHOLDER_SUBSECTION) ? 1 : 0);
				}
				else
				{
					num = 1;
				}
			}
			else
			{
				num = 0;
			}
			flag = (byte)num != 0;
			textFrame = null;
		}
		else if (shp.Type == MsoShapeType.msoGroup)
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = shp.GroupItems.GetEnumerator();
				while (true)
				{
					if (enumerator.MoveNext())
					{
						flag = HasSectionTitlePlaceholder((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current);
						if (!flag)
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
								goto end_IL_00c0;
							}
							continue;
							end_IL_00c0:
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
							goto end_IL_00d6;
						}
						continue;
						end_IL_00d6:
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
		}
		return flag;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = A;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = presentation.Slides.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				Slide slide = (Slide)enumerator.Current;
				SlideType slideType = PowerPointAddIn1.Slides.Helpers.GetSlideType(slide);
				if ((uint)(slideType - 1) <= 4u)
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
				if ((uint)(slideType - 7) <= 3u)
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
					enumerator2 = slide.CustomLayout.Shapes.GetEnumerator();
					while (true)
					{
						if (enumerator2.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
							try
							{
								if (IsLegacySectionTitlePlaceholder(shape))
								{
									B(slide, shape);
									shape.TextFrame2.TextRange.Text = "";
									shape.Visible = MsoTriState.msoFalse;
									break;
								}
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
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_00cf;
							}
							continue;
							end_IL_00cf:
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
							switch (6)
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
			while (true)
			{
				switch (4)
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
		presentation.Tags.Delete(SectionTitles.m_B);
		presentation = null;
	}

	private static void A(bool A)
	{
		SectionTitles.A(A, null);
	}

	private static void A(bool A, Slide B = null)
	{
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = NG.A.Application.ActivePresentation;
		string text = "";
		SectionProperties sectionProperties = activePresentation.SectionProperties;
		checked
		{
			bool flag2 = default(bool);
			if (B == null)
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
				int count = sectionProperties.Count;
				IEnumerator enumerator = default(IEnumerator);
				for (int i = 2; i <= count; i++)
				{
					string text2 = sectionProperties.Name(i);
					string d;
					if (!PowerPointAddIn1.Agenda.Update.A(text2))
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
						text = text2;
						d = "";
					}
					else
					{
						d = Strings.Mid(text2, 2);
					}
					int num2;
					int num = (num2 = sectionProperties.FirstSlide(i)) + sectionProperties.SlidesCount(i) - 1;
					for (int j = num2; j <= num; j++)
					{
						Slide slide = activePresentation.Slides[j];
						SlideType slideType = PowerPointAddIn1.Slides.Helpers.GetSlideType(slide);
						unchecked
						{
							if ((uint)(slideType - 1) > 4u)
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
								if ((uint)(slideType - 7) > 3u)
								{
									bool flag = false;
									try
									{
										enumerator = slide.CustomLayout.Shapes.GetEnumerator();
										while (enumerator.MoveNext())
										{
											Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
											try
											{
												if (!IsLegacySectionTitlePlaceholder(shape))
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
													SectionTitles.A(slide, shape, text, d);
													flag = true;
													break;
												}
											}
											catch (Exception ex)
											{
												ProjectData.SetProjectError(ex);
												Exception ex2 = ex;
												ProjectData.ClearProjectError();
											}
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
									if (!flag)
									{
										flag2 = true;
									}
								}
							}
							slide = null;
						}
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_018a;
						}
						continue;
						end_IL_018a:
						break;
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
			else
			{
				int sectionIndex = B.sectionIndex;
				string text2 = sectionProperties.Name(sectionIndex);
				string d;
				if (!PowerPointAddIn1.Agenda.Update.A(text2))
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
					text = text2;
					d = "";
				}
				else
				{
					d = Strings.Mid(text2, 2);
					for (int k = sectionIndex - 1; k >= 2; k += -1)
					{
						if (PowerPointAddIn1.Agenda.Update.A(sectionProperties.Name(k)))
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
						text = sectionProperties.Name(k);
						break;
					}
				}
				if (Operators.CompareString(text, AH.A(148954), TextCompare: false) != 0 && Operators.CompareString(text, AH.A(7090), TextCompare: false) != 0)
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
					IEnumerator enumerator2 = default(IEnumerator);
					try
					{
						enumerator2 = B.CustomLayout.Shapes.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
							try
							{
								if (IsLegacySectionTitlePlaceholder(shape2))
								{
									SectionTitles.A(B, shape2, text, d);
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
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (2)
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
			}
			sectionProperties = null;
			activePresentation.Tags.Add(SectionTitles.m_B, AH.A(9078));
			if (!A)
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
				if (flag2)
				{
					Forms.WarningMessage(AH.A(149264));
				}
			}
			activePresentation = null;
		}
	}

	private static void A(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B, string C, string D)
	{
		SectionTitles.B(A, B);
		Microsoft.Office.Interop.PowerPoint.Shape shape = B;
		shape.Visible = MsoTriState.msoTrue;
		clsClipboard.CopyWithWait((Action)([SpecialName] () =>
		{
			B.Copy();
		}), 4000);
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = AddRemove.TryPaste(B, A.Shapes);
		TextRange textRange = shape2.TextFrame.TextRange;
		string text;
		if (D.Length != 0)
		{
			if (C.Length != 0)
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
				text = C + KG.A.SectionSubsectionSeparator + D;
			}
			else
			{
				text = D;
			}
		}
		else
		{
			text = C;
		}
		textRange.Text = text;
		shape2.Visible = MsoTriState.msoTrue;
		_ = null;
		shape.Visible = MsoTriState.msoFalse;
		_ = null;
	}

	private static void B(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		float top = B.Top;
		float left = B.Left;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Shapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				if (shape.Top != top)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (shape.Left == left)
				{
					shape.Delete();
					break;
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

	public static bool IsLegacySectionTitlePlaceholder(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		return PowerPointAddIn1.Shapes.Helpers.IsShapeType(shp, AH.A(149450));
	}

	public static void ConvertLegacySectionTitles()
	{
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
		bool flag = false;
		bool flag2 = false;
		bool flag3 = false;
		bool flag4 = false;
		try
		{
			presentation = NG.A.Application.ActivePresentation;
			flag = Conversions.ToBoolean(presentation.Tags[SectionTitles.m_A]);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
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
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Master slideMaster;
			Microsoft.Office.Interop.PowerPoint.Shape shape;
			List<int> list;
			try
			{
				slideMaster = presentation.Designs[1].SlideMaster;
				IEnumerator enumerator = slideMaster.Shapes.GetEnumerator();
				try
				{
					while (true)
					{
						if (enumerator.MoveNext())
						{
							if (!A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current))
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
								flag = true;
								break;
							}
							break;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_00c1;
							}
							continue;
							end_IL_00c1:
							break;
						}
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
				if (!flag)
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
					IEnumerator enumerator2 = default(IEnumerator);
					try
					{
						enumerator2 = slideMaster.CustomLayouts.GetEnumerator();
						IEnumerator enumerator3 = default(IEnumerator);
						while (true)
						{
							if (enumerator2.MoveNext())
							{
								CustomLayout customLayout = (CustomLayout)enumerator2.Current;
								list = new List<int>();
								shape = null;
								for (int i = customLayout.Shapes.Count; i >= 1; i = checked(i + -1))
								{
									try
									{
										if (!IsLegacySectionTitlePlaceholder(customLayout.Shapes[i]))
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
											list.Add(PowerPointAddIn1.Shapes.Helpers.A(customLayout, customLayout.Shapes[i]));
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
									switch (1)
									{
									case 0:
										continue;
									}
									break;
								}
								if (list.Any())
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
									if (!flag3)
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
										if (MessageBox.Show(AH.A(149465), AH.A(5874), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
										{
											while (true)
											{
												switch (3)
												{
												case 0:
													continue;
												}
												flag4 = true;
												break;
											}
											break;
										}
										A(presentation);
										flag3 = true;
									}
									Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = customLayout.Shapes.Range(list.ToArray());
									if (list.Count == 1)
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
										shape = shapeRange[1];
									}
									else
									{
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
											{
												enumerator3 = shapeRange.GetEnumerator();
												try
												{
													while (enumerator3.MoveNext())
													{
														((Microsoft.Office.Interop.PowerPoint.Shape)enumerator3.Current).Visible = MsoTriState.msoTrue;
													}
													while (true)
													{
														switch (1)
														{
														case 0:
															break;
														default:
															goto end_IL_0281;
														}
														continue;
														end_IL_0281:
														break;
													}
												}
												finally
												{
													IDisposable disposable2 = enumerator3 as IDisposable;
													if (disposable2 != null)
													{
														disposable2.Dispose();
													}
												}
											}
										}
										shape = shapeRange.Group();
									}
									shapeRange = null;
								}
								if (shape == null)
								{
									continue;
								}
								if (!flag2)
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
									shape.Copy();
									Microsoft.Office.Interop.PowerPoint.Shape shape2 = slideMaster.Shapes.Paste()[1];
									shape2.Visible = MsoTriState.msoFalse;
									shape2.Name = AH.A(149631);
									shape2.TextFrame2.TextRange.Text = Placeholders.PLACEHOLDER_SECTION + KG.A.SectionSubsectionSeparator + Placeholders.PLACEHOLDER_SUBSECTION;
									shape2.Top = shape.Top;
									shape2.Left = shape.Left;
									_ = null;
									flag2 = true;
								}
								shape.Delete();
								continue;
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_0359;
								}
								continue;
								end_IL_0359:
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
								switch (6)
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
				if (!flag4)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						presentation.Tags.Add(SectionTitles.m_A, AH.A(149670));
						break;
					}
				}
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			slideMaster = null;
			shape = null;
			list = null;
		}
		presentation = null;
	}

	private static bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		bool flag = false;
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		if (shape.Type != MsoShapeType.msoPlaceholder)
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
			if (shape.HasTextFrame == MsoTriState.msoTrue)
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
				if (shape.TextFrame2.HasText == MsoTriState.msoTrue)
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
					if (shape.TextFrame2.TextRange.Text.Contains(Placeholders.PLACEHOLDER_SECTION))
					{
						flag = true;
					}
				}
			}
			else if (shape.Type == MsoShapeType.msoGroup)
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = shape.GroupItems.GetEnumerator();
					while (true)
					{
						if (enumerator.MoveNext())
						{
							flag = SectionTitles.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current);
							if (!flag)
							{
								continue;
							}
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
									goto end_IL_00b7;
								}
								continue;
								end_IL_00b7:
								break;
							}
							break;
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_00cd;
							}
							continue;
							end_IL_00cd:
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
		shape = null;
		return flag;
	}
}
