using System.Reflection;
using System.Runtime.CompilerServices;
using MacabacusMacros.ExcelHelpers;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Shapes;

public sealed class Navigate
{
	internal static void A(Shape A)
	{
		Shape shape = A;
		Ranges.ScrollIntoView(((_Application)shape.Application).get_Range((object)shape.TopLeftCell, (object)shape.BottomRightCell));
		if (shape.Visible == MsoTriState.msoTrue)
		{
			shape.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		shape = null;
	}
}
