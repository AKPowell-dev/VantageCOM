using System;
using System.Collections;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1;

public sealed class clsTest
{
	public static void CheckShapes()
	{
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = NG.A.Application.ActivePresentation;
		float num = clsPublish.InchesToPoints(1f);
		string text = "";
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = activePresentation.Slides.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				Slide slide = (Slide)enumerator.Current;
				if (slide.Shapes.Count <= 2)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				try
				{
					enumerator2 = slide.Shapes.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
						if (Math.Round(shape.Height, 4) != (double)num)
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
							if (Math.Round(shape.Width, 4) != (double)num)
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
								Interaction.MsgBox(Conversions.ToString(shape.Height) + AH.A(105620) + Conversions.ToString(shape.Width));
								text = text + AH.A(33556) + shape.Name + AH.A(105627) + Conversions.ToString(slide.SlideIndex) + AH.A(105650);
								continue;
							}
						}
						if (shape.Type != MsoShapeType.msoFreeform && shape.Type != MsoShapeType.msoGroup)
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
							text = text + AH.A(33556) + shape.Name + AH.A(105627) + Conversions.ToString(slide.SlideIndex) + AH.A(105739);
						}
						else if (shape.LockAspectRatio != MsoTriState.msoTrue)
						{
							text = text + AH.A(33556) + shape.Name + AH.A(105627) + Conversions.ToString(slide.SlideIndex) + AH.A(105810);
						}
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_0267;
						}
						continue;
						end_IL_0267:
						break;
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
		if (text.Length > 0)
		{
			Clipboard.SetText(text);
			MessageBox.Show(AH.A(105889));
		}
		else
		{
			MessageBox.Show(AH.A(105944));
		}
		activePresentation = null;
	}
}
