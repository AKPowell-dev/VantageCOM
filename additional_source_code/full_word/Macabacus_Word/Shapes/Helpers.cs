using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Shapes;

public sealed class Helpers
{
	public static Microsoft.Office.Interop.Word.ShapeRange SelectedShapes(Selection sel)
	{
		if (A(sel))
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
					return sel.ChildShapeRange;
				}
			}
		}
		return sel.ShapeRange;
	}

	private static bool A(Selection A)
	{
		if (A.ChildShapeRange.Count > 0)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Microsoft.Office.Interop.Word.ShapeRange childShapeRange = A.ChildShapeRange;
					object Index = 1;
					return childShapeRange[ref Index].Type != MsoShapeType.msoGroup;
				}
				}
			}
		}
		return false;
	}

	internal static bool A(Range A)
	{
		bool result;
		try
		{
			_ = A.ShapeRange;
			result = true;
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
}
