using System;
using System.Collections;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes;

public sealed class Duplicate
{
	public static void Shape()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Selection selection;
			try
			{
				selection = application.ActiveWindow.Selection;
				if (selection.Type == PpSelectionType.ppSelectionShapes)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						if (selection.ShapeRange.Count == 1)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								Shape shape = selection.ShapeRange[1];
								float left = shape.Left;
								float top = shape.Top;
								shape.Copy();
								int slideIndex = selection.SlideRange[1].SlideIndex;
								application.StartNewUndoEntry();
								try
								{
									enumerator = application.ActivePresentation.Slides.GetEnumerator();
									while (enumerator.MoveNext())
									{
										Slide slide = (Slide)enumerator.Current;
										if (slide.SlideIndex != slideIndex)
										{
											Shape shape2 = slide.Shapes.Paste()[1];
											shape2.Top = top;
											shape2.Left = left;
											_ = null;
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
								break;
							}
						}
						else
						{
							Helpers.SingleShapeRequiredError();
						}
						break;
					}
				}
				else
				{
					Forms.WarningMessage(AH.A(73308));
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.WarningMessage(AH.A(73308));
				ProjectData.ClearProjectError();
			}
			application = null;
			selection = null;
			return;
		}
	}
}
