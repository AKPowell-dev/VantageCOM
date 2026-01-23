using System;
using System.Collections;
using System.Drawing;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Pagination;

public sealed class FacingSlides
{
	public static readonly string TAG_NAME = AH.A(101546);

	public static void MarkSlide()
	{
		if (!Licensing.AllowPaginationOperation())
		{
			return;
		}
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
			Application application;
			Selection selection;
			Slide slide;
			try
			{
				application = NG.A.Application;
				selection = application.ActiveWindow.Selection;
				if (selection.SlideRange.Count == 1)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						slide = selection.SlideRange[1];
						SlideType slideType = Helpers.GetSlideType(slide);
						if (slideType != SlideType.Flysheet)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								if (slideType != SlideType.Blank)
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											continue;
										}
										A(slide, application.ActivePresentation);
										clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)4, AH.A(101344));
										break;
									}
								}
								else
								{
									A(AH.A(101255));
								}
								break;
							}
						}
						else
						{
							A(AH.A(101172));
						}
						break;
					}
				}
				else
				{
					A(AH.A(101385));
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				A(AH.A(101385));
				ProjectData.ClearProjectError();
			}
			application = null;
			selection = null;
			slide = null;
			return;
		}
	}

	internal static void A(Slide A, Microsoft.Office.Interop.PowerPoint.Presentation B)
	{
		A.Application.StartNewUndoEntry();
		Microsoft.Office.Interop.PowerPoint.Shape shape = A.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, -25f, 0f, 20f, B.PageSetup.SlideHeight);
		shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.IndianRed);
		shape.Line.Visible = MsoTriState.msoFalse;
		Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = shape.TextFrame2;
		textFrame.Orientation = MsoTextOrientation.msoTextOrientationHorizontalRotatedFarEast;
		TextRange2 textRange = textFrame.TextRange;
		textRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
		textRange.Text = AH.A(101482);
		Font2 font = textRange.Font;
		font.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
		font.Bold = MsoTriState.msoTrue;
		_ = null;
		_ = null;
		_ = null;
		shape.Tags.Add(TAG_NAME, AH.A(9078));
		_ = null;
	}

	public static void UnmarkSlide()
	{
		if (!Licensing.AllowPaginationOperation())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Application application;
			Selection selection;
			try
			{
				application = NG.A.Application;
				selection = application.ActiveWindow.Selection;
				if (selection.SlideRange.Count > 0)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						application.StartNewUndoEntry();
						try
						{
							enumerator = selection.SlideRange.GetEnumerator();
							while (enumerator.MoveNext())
							{
								A((Slide)enumerator.Current);
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_0087;
								}
								continue;
								end_IL_0087:
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
						break;
					}
				}
				else
				{
					A(AH.A(101507));
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				A(AH.A(101507));
				ProjectData.ClearProjectError();
			}
			application = null;
			selection = null;
			return;
		}
	}

	internal static void A(Slide A)
	{
		for (int i = A.Shapes.Count; i >= 1; i = checked(i + -1))
		{
			try
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = A.Shapes[i];
				if (Operators.CompareString(shape.Tags[TAG_NAME], AH.A(9078), TextCompare: false) == 0)
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
						shape.Delete();
						return;
					}
				}
				shape = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
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

	private static void A(string A)
	{
		Forms.WarningMessage(A);
	}

	public static bool IsFacingSlide(Slide sld)
	{
		try
		{
			for (int i = sld.Shapes.Count; i >= 1; i = checked(i + -1))
			{
				if (Operators.CompareString(sld.Shapes[i].Tags[TAG_NAME], AH.A(9078), TextCompare: false) == 0)
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
							return true;
						}
					}
				}
				_ = null;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_006b;
				}
				continue;
				end_IL_006b:
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return false;
	}
}
