using System;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes;

public sealed class Split
{
	public static void Shape()
	{
		if (!Licensing.AllowAdvancedShapeOperation())
		{
			return;
		}
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
			ShapeRange shapeRange;
			try
			{
				shapeRange = Base.SelectedShapes();
				if (shapeRange.Count != 1)
				{
					throw new Exception();
				}
				NG.A.Application.StartNewUndoEntry();
				new wpfSplitShape(shapeRange[1]).Show();
				_ = null;
				Base.LogActivity(AH.A(75062));
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
