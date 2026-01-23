using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Libraries.Pane;
using MacabacusMacros.Libraries.Versioning;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.Library2.Insert;

namespace PowerPointAddIn1.Library2.Versioning.Replace;

public sealed class Shapes
{
	[CompilerGenerated]
	internal sealed class KD
	{
		public Shape A;

		[SpecialName]
		internal void A()
		{
			this.A.Copy();
		}
	}

	internal static void A(ShapeItem A, Application B, ref List<string> C)
	{
		//IL_00e8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f8: Unknown result type (might be due to invalid IL or missing references)
		//IL_0100: Unknown result type (might be due to invalid IL or missing references)
		Shape A2;
		try
		{
			Microsoft.Office.Interop.PowerPoint.Presentation presentation = PowerPointAddIn1.Library2.Insert.Common.A(Common.A((ContentItem)(object)A), B, ref C);
			Shape shape = A.Shape;
			string value = Tagging.A(shape.Tags);
			if (B.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionText)
			{
				shape.Select();
			}
			if (A.A())
			{
				while (true)
				{
					switch (5)
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
				HD.A(presentation);
				goto IL_012f;
			}
			if (presentation.Slides.Count == 1)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					break;
				}
				if (presentation.Slides[1].Shapes.Count == 1)
				{
					A2 = presentation.Slides[1].Shapes[1];
					goto IL_0116;
				}
			}
			A2 = PowerPointApp.GetSourceShape(presentation, ((ContentItem)A).ManifestInfo.SlideIndex, ((ContentItem)A).ContentInfo.ContentId, ((ContentItem)A).ContentInfo.Title);
			goto IL_0116;
			IL_0116:
			clsClipboard.CopyWithWait((Action)([SpecialName] () =>
			{
				A2.Copy();
			}), 4000);
			goto IL_012f;
			IL_012f:
			Shape shape2 = PowerPointAddIn1.Library2.Insert.Shapes.A(B);
			Shape shape3 = shape;
			shape2.Top = shape3.Top;
			shape2.Left = shape3.Left;
			shape2.Height = shape3.Height;
			shape2.Width = shape3.Width;
			shape3.Delete();
			shape3 = null;
			shape2.Select();
			shape2.Tags.Add(Tagging.A, value);
			A.Shape = shape2;
		}
		finally
		{
			Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
			A2 = null;
			Shape shape2 = null;
			Shape shape = null;
		}
		clsClipboard.ClearClipboard();
	}
}
