using System;
using System.Windows.Forms;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Library2.Insert;

public sealed class Shapes
{
	internal static Shape A(Microsoft.Office.Interop.PowerPoint.Application A)
	{
		Shape result;
		try
		{
			A.CommandBars.ExecuteMso(AH.A(58900));
			System.Windows.Forms.Application.DoEvents();
			ShapeRange shapeRange = A.ActiveWindow.Selection.ShapeRange;
			Shape shape;
			if (shapeRange.Count != 1)
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
				shape = shapeRange.Group();
			}
			else
			{
				shape = shapeRange[1];
			}
			result = shape;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ShapeRange shapeRange2 = A.ActiveWindow.Selection.SlideRange[1].Shapes.Paste();
			Shape shape2;
			if (shapeRange2.Count != 1)
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
				shape2 = shapeRange2.Group();
			}
			else
			{
				shape2 = shapeRange2[1];
			}
			result = shape2;
			ProjectData.ClearProjectError();
		}
		return result;
	}
}
