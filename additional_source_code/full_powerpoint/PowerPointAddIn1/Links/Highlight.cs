using System;
using System.Collections;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using A;
using MacabacusMacros.Links;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Links;

public sealed class Highlight
{
	private static int m_A;

	private static readonly string m_A = AH.A(94243);

	private static readonly int B = 6;

	[CompilerGenerated]
	private static bool m_A;

	[CompilerGenerated]
	private static PpViewType m_A;

	public static bool IsHighlighted
	{
		[CompilerGenerated]
		get
		{
			return Highlight.m_A;
		}
		[CompilerGenerated]
		set
		{
			Highlight.m_A = value;
		}
	} = false;

	private static PpViewType OriginalViewType
	{
		[CompilerGenerated]
		get
		{
			return Highlight.m_A;
		}
		[CompilerGenerated]
		set
		{
			Highlight.m_A = value;
		}
	}

	public static void LoadColor(int i)
	{
		Color c;
		if (i != 0)
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
			if (i != 1)
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
				c = Color.FromArgb(0, 255, 0);
			}
			else
			{
				c = Color.FromArgb(255, 255, 0);
			}
		}
		else
		{
			c = Color.FromArgb(0, 255, 255);
		}
		Highlight.m_A = ColorTranslator.ToOle(c);
	}

	public static void Toggle(bool blnPressed)
	{
		if (blnPressed)
		{
			A();
		}
		else
		{
			Remove();
		}
	}

	private static void A()
	{
		Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
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
				OriginalViewType = application.ActiveWindow.ViewType;
				application.ActiveWindow.ViewType = PpViewType.ppViewSlideSorter;
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = presentation.Slides.GetEnumerator();
					IEnumerator enumerator2 = default(IEnumerator);
					IEnumerator enumerator3 = default(IEnumerator);
					while (enumerator.MoveNext())
					{
						Slide slide = (Slide)enumerator.Current;
						try
						{
							enumerator2 = slide.Shapes.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
								if (shape.Visible == MsoTriState.msoTrue)
								{
									A(shape, slide);
								}
							}
							while (true)
							{
								switch (6)
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
						try
						{
							enumerator3 = slide.Hyperlinks.GetEnumerator();
							while (enumerator3.MoveNext())
							{
								Hyperlink hyp = (Hyperlink)enumerator3.Current;
								if (!Hyperlinks.IsLinked(hyp))
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
								Microsoft.Office.Interop.PowerPoint.Shape parentShape = Hyperlinks.GetParentShape(hyp, blnIgnoreTables: true);
								float num = 0f;
								float num2 = 0f;
								if (Hyperlinks.IsShapeInsideTableCell(parentShape))
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
									num = parentShape.Left;
									num2 = parentShape.Top;
								}
								TextRange2 textRange = Hyperlinks.HyperlinkParentTextRange2(hyp);
								A(slide, textRange.BoundLeft - (float)B + num, textRange.BoundTop - (float)B + num2, textRange.BoundWidth + (float)(B * 2), textRange.BoundHeight + (float)(B * 2));
								textRange = null;
								parentShape = null;
							}
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
									goto end_IL_01e4;
								}
								continue;
								end_IL_01e4:
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
				IsHighlighted = true;
				new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(application, new EApplication_WindowSelectionChangeEventHandler(A));
				presentation = null;
			}
			Common.LogActivity(AH.A(94181));
			application = null;
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, Slide B)
	{
		checked
		{
			if (A.Type != MsoShapeType.msoGroup)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (Shapes.IsLinked(A))
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									Highlight.A(B, A.Left - (float)Highlight.B, A.Top - (float)Highlight.B, A.Width + (float)(Highlight.B * 2), A.Height + (float)(Highlight.B * 2));
									return;
								}
							}
						}
						if (Text.ContainsLinks(A))
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
								{
									foreach (TextLink item in Text.SelectedLinks(A))
									{
										TextRange2 textRange = item.TextRange;
										Highlight.A(B, textRange.BoundLeft - (float)Highlight.B, textRange.BoundTop - (float)Highlight.B, textRange.BoundWidth + (float)(Highlight.B * 2), textRange.BoundHeight + (float)(Highlight.B * 2));
										textRange = null;
									}
									return;
								}
								}
							}
						}
						return;
					}
				}
			}
			IEnumerator enumerator2 = A.GroupItems.GetEnumerator();
			try
			{
				while (enumerator2.MoveNext())
				{
					Highlight.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, B);
				}
				while (true)
				{
					switch (2)
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
				IDisposable disposable = enumerator2 as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
		}
	}

	private static void A(Slide A, float B, float C, float D, float E)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = A.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, B, C, D, E);
		shape.Name = Highlight.m_A;
		Microsoft.Office.Interop.PowerPoint.LineFormat line = shape.Line;
		line.Transparency = 0.25f;
		line.Weight = Highlight.B;
		line.Style = MsoLineStyle.msoLineSingle;
		line.DashStyle = MsoLineDashStyle.msoLineSolid;
		line.ForeColor.RGB = Highlight.m_A;
		_ = null;
		shape.Fill.ForeColor.RGB = Highlight.m_A;
		shape.Fill.Transparency = 0.75f;
		_ = null;
	}

	public static void Remove()
	{
		Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
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
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = presentation.Slides.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shapes shapes = ((Slide)enumerator.Current).Shapes;
					for (int i = shapes.Count; i >= 1; i = checked(i + -1))
					{
						if (Operators.CompareString(shapes[i].Name, Highlight.m_A, TextCompare: false) != 0)
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
						shapes[i].Delete();
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_00b9;
						}
						continue;
						end_IL_00b9:
						break;
					}
					shapes = null;
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_00d3;
					}
					continue;
					end_IL_00d3:
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
			IsHighlighted = false;
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(application, new EApplication_WindowSelectionChangeEventHandler(A));
			if (OriginalViewType != PpViewType.ppViewSlideSorter)
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
				application.ActiveWindow.ViewType = PpViewType.ppViewNormal;
			}
			presentation = null;
		}
		application = null;
	}

	private static void A(Selection A)
	{
		if (A.Application.ActiveWindow.ViewType == PpViewType.ppViewSlideSorter)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Remove();
			KG.A.InvalidateControl(AH.A(94212));
			return;
		}
	}
}
