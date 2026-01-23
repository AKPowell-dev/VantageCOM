using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using MacabacusMacros;
using MacabacusMacros.Libraries.Pane;
using MacabacusMacros.Libraries.Versioning;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Library2.Versioning.Replace;

public sealed class Shapes
{
	[CompilerGenerated]
	internal sealed class R
	{
		public Microsoft.Office.Interop.PowerPoint.Shape A;

		[SpecialName]
		internal void A()
		{
			this.A.Copy();
		}
	}

	internal static void A(ShapeItem A, Microsoft.Office.Interop.Word.Application B, ref Microsoft.Office.Interop.PowerPoint.Application C, ref bool D, ref List<string> E)
	{
		//IL_0099: Unknown result type (might be due to invalid IL or missing references)
		//IL_009e: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b8: Unknown result type (might be due to invalid IL or missing references)
		string text = Common.A((ContentItem)(object)A);
		Presentation presentation = PowerPointApp.GetPresentation(ref C, ref E, ref D, text);
		Microsoft.Office.Interop.PowerPoint.Shape A2;
		if (presentation.Slides.Count == 1)
		{
			while (true)
			{
				switch (3)
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
			if (presentation.Slides[1].Shapes.Count == 1)
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
				A2 = presentation.Slides[1].Shapes[1];
				goto IL_00cb;
			}
		}
		A2 = PowerPointApp.GetSourceShape(presentation, ((ContentItem)A).ManifestInfo.SlideIndex, ((ContentItem)A).ContentInfo.ContentId, ((ContentItem)A).ContentInfo.Title);
		goto IL_00cb;
		IL_00cb:
		clsClipboard.CopyWithWait((Action)([SpecialName] () =>
		{
			A2.Copy();
		}), 4000);
		if (A.Shape is Microsoft.Office.Interop.Word.Shape)
		{
			Microsoft.Office.Interop.Word.Shape shape = (Microsoft.Office.Interop.Word.Shape)A.Shape;
			Common.Q b = Common.A(shape);
			try
			{
				string alternativeText = shape.AlternativeText;
				B.Selection.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting);
				Microsoft.Office.Interop.Word.ShapeRange shapeRange = B.Selection.ShapeRange;
				object Index = 1;
				Microsoft.Office.Interop.Word.Shape shape2 = shapeRange[ref Index];
				Common.A(shape2, b);
				shape.Delete();
				Microsoft.Office.Interop.Word.Shape shape3 = shape2;
				Index = RuntimeHelpers.GetObjectValue(Missing.Value);
				shape3.Select(ref Index);
				shape2.AlternativeText = alternativeText;
				A.Shape = shape2;
			}
			finally
			{
				shape = null;
				Microsoft.Office.Interop.Word.Shape shape2 = null;
				presentation = null;
				A2 = null;
			}
			clsClipboard.ClearClipboard();
			return;
		}
		throw new NotImplementedException();
	}
}
