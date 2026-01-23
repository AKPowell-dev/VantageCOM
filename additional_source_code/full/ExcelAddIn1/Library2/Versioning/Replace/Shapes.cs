using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using ExcelAddIn1.Library2.Insert;
using MacabacusMacros;
using MacabacusMacros.Libraries.Pane;
using MacabacusMacros.Libraries.Versioning;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;

namespace ExcelAddIn1.Library2.Versioning.Replace;

public sealed class Shapes
{
	internal static void A(ShapeItem A, Microsoft.Office.Interop.Excel.Application B, ref Microsoft.Office.Interop.PowerPoint.Application C, ref bool D, ref List<string> E)
	{
		//IL_009d: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00aa: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ba: Unknown result type (might be due to invalid IL or missing references)
		Microsoft.Office.Interop.Excel.Shape shape = A.Shape;
		string text = Common.A((ContentItem)(object)A);
		Worksheet b = (Worksheet)B.ActiveSheet;
		Presentation presentation = PowerPointApp.GetPresentation(ref C, ref E, ref D, text);
		Microsoft.Office.Interop.PowerPoint.Shape a;
		if (presentation.Slides.Count == 1 && presentation.Slides[1].Shapes.Count == 1)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			a = presentation.Slides[1].Shapes[1];
		}
		else
		{
			a = PowerPointApp.GetSourceShape(presentation, ((ContentItem)A).ManifestInfo.SlideIndex, ((ContentItem)A).ContentInfo.ContentId, ((ContentItem)A).ContentInfo.Title);
		}
		try
		{
			Microsoft.Office.Interop.Excel.Shape shape2 = ExcelAddIn1.Library2.Insert.Shapes.A(a, b);
			Microsoft.Office.Interop.Excel.Shape shape3 = shape;
			shape2.AlternativeText = shape3.AlternativeText;
			shape2.Width = shape3.Width;
			shape2.Height = shape3.Height;
			shape2.Top = shape3.Top;
			shape2.Left = shape3.Left;
			shape3.Delete();
			shape3 = null;
			shape2.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
			A.Shape = shape2;
			clsClipboard.ClearClipboard();
		}
		finally
		{
			shape = null;
			Microsoft.Office.Interop.Excel.Shape shape2 = null;
			presentation = null;
			a = null;
		}
	}
}
