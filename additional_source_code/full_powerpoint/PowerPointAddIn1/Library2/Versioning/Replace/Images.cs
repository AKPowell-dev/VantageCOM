using System;
using System.Collections.Generic;
using MacabacusMacros;
using MacabacusMacros.Libraries.Versioning;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Library2.Insert;

namespace PowerPointAddIn1.Library2.Versioning.Replace;

public sealed class Images
{
	internal static void A(ShapeItem A, Application B)
	{
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = B.ActivePresentation;
		bool flag = false;
		try
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape = A.Shape;
			Slide slideFromShape = clsPowerPoint.GetSlideFromShape(shape);
			try
			{
				flag = PowerPointAddIn1.Library2.Insert.Images.A(A.Shape);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (!flag)
			{
				PageSetup pageSetup = activePresentation.PageSetup;
				_ = pageSetup.SlideWidth;
				_ = pageSetup.SlideHeight;
				_ = null;
			}
			string value = Tagging.A(shape.Tags);
			string text = Common.A((ContentItem)(object)A);
			Microsoft.Office.Interop.PowerPoint.Shape shape2 = default(Microsoft.Office.Interop.PowerPoint.Shape);
			if (flag)
			{
				int B2 = 0;
				Dictionary<Microsoft.Office.Interop.PowerPoint.Shape, PowerPointAddIn1.Library2.Insert.Images.FD> dictionary = new Dictionary<Microsoft.Office.Interop.PowerPoint.Shape, PowerPointAddIn1.Library2.Insert.Images.FD>();
				Dictionary<int, PowerPointAddIn1.Library2.Insert.Images.FD> C = new Dictionary<int, PowerPointAddIn1.Library2.Insert.Images.FD>();
				shape = PowerPointAddIn1.Library2.Insert.Images.A(shape, slideFromShape);
				PowerPointAddIn1.Library2.Insert.Images.A(slideFromShape, ref B2, ref C);
				if (B2 > 0)
				{
					int zOrderPosition = shape.ZOrderPosition;
					int num = B2;
					int num2 = 1;
					while (true)
					{
						if (num2 <= num)
						{
							shape2 = PowerPointAddIn1.Library2.Insert.Images.A(slideFromShape, text);
							if (shape2.ZOrderPosition != zOrderPosition)
							{
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
									dictionary.Add(shape2, C[shape2.ZOrderPosition]);
									num2 = checked(num2 + 1);
									break;
								}
								continue;
							}
							PowerPointAddIn1.Library2.Insert.Images.A(shape2, B);
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
						break;
					}
					PowerPointAddIn1.Library2.Insert.Images.A(slideFromShape, dictionary);
				}
				dictionary = null;
				C = null;
			}
			else
			{
				shape2 = PowerPointAddIn1.Library2.Insert.Images.A(B, slideFromShape, text);
				shape2.LockAspectRatio = MsoTriState.msoTrue;
				Microsoft.Office.Interop.PowerPoint.Shape shape3 = shape;
				shape2.Top = shape3.Top;
				shape2.Left = shape3.Left;
				shape2.Width = shape3.Width;
				PowerPointAddIn1.Library2.Insert.Images.A(shape2, shape3.ZOrderPosition);
				shape3 = null;
				shape.Delete();
				shape2.Select();
			}
			shape2.Tags.Add(Tagging.A, value);
			A.Shape = shape2;
		}
		finally
		{
			Slide slideFromShape = null;
			Microsoft.Office.Interop.PowerPoint.Shape shape = null;
			Microsoft.Office.Interop.PowerPoint.Shape shape2 = null;
			activePresentation = null;
		}
	}
}
