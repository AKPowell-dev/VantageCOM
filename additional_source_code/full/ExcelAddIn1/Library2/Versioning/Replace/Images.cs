using System.Reflection;
using System.Runtime.CompilerServices;
using MacabacusMacros.Libraries.Versioning;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Library2.Versioning.Replace;

public sealed class Images
{
	internal static void A(ShapeItem A, Application B)
	{
		Shape shape = A.Shape;
		string alternativeText = shape.AlternativeText;
		string filename = Common.A((ContentItem)(object)A);
		try
		{
			Shape shape2 = ((Worksheet)B.ActiveSheet).Shapes.AddPicture2(filename, MsoTriState.msoFalse, MsoTriState.msoTrue, shape.Left, shape.Top, -1f, -1f, MsoPictureCompress.msoPictureCompressDocDefault);
			shape2.LockAspectRatio = MsoTriState.msoTrue;
			shape2.Width = shape.Width;
			shape.Delete();
			shape2.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
			shape2.AlternativeText = alternativeText;
			A.Shape = shape2;
		}
		finally
		{
			Shape shape2 = null;
			shape = null;
		}
	}
}
