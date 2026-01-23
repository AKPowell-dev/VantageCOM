using System;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes;

public sealed class Multiply
{
	public static void Shape()
	{
		if (!Licensing.AllowAdvancedShapeOperation())
		{
			return;
		}
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
			ShapeRange shapeRange;
			try
			{
				shapeRange = Base.SelectedShapes();
				if (shapeRange.Count != 1)
				{
					throw new Exception();
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					NG.A.Application.StartNewUndoEntry();
					new wpfMultiplyShape(shapeRange[1]).Show();
					_ = null;
					Base.LogActivity(AH.A(83086));
					break;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Helpers.SingleShapeRequiredError();
				ProjectData.ClearProjectError();
			}
			shapeRange = null;
			return;
		}
	}
}
