using System.Reflection;
using System.Runtime.CompilerServices;
using MacabacusMacros.Libraries.Versioning;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Library2.Versioning.Replace;

public sealed class Images
{
	internal static void A(ShapeItem A, Application B)
	{
		string fileName = Common.A((ContentItem)(object)A);
		if (A.Shape is InlineShape)
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
					InlineShape inlineShape = (InlineShape)A.Shape;
					try
					{
						string alternativeText = inlineShape.AlternativeText;
						InlineShapes inlineShapes = B.Selection.Range.InlineShapes;
						object LinkToFile = false;
						object SaveWithDocument = true;
						object Range = inlineShape.Range;
						InlineShape inlineShape2 = inlineShapes.AddPicture(fileName, ref LinkToFile, ref SaveWithDocument, ref Range);
						inlineShape2.LockAspectRatio = MsoTriState.msoTrue;
						inlineShape2.Width = inlineShape.Width;
						inlineShape.Delete();
						inlineShape2.Select();
						inlineShape2.AlternativeText = alternativeText;
						A.Shape = inlineShape2;
						return;
					}
					finally
					{
						inlineShape = null;
						InlineShape inlineShape2 = null;
					}
				}
				}
			}
		}
		Microsoft.Office.Interop.Word.Shape shape = (Microsoft.Office.Interop.Word.Shape)A.Shape;
		Common.Q b = Common.A(shape);
		try
		{
			string alternativeText = shape.AlternativeText;
			Microsoft.Office.Interop.Word.Shape shape2 = shape;
			Microsoft.Office.Interop.Word.Shapes shapes = B.ActiveDocument.Shapes;
			object Range = false;
			object SaveWithDocument = true;
			Microsoft.Office.Interop.Word.Shape shape3;
			object LinkToFile = (shape3 = shape2).Left;
			Microsoft.Office.Interop.Word.Shape shape4;
			object Top = (shape4 = shape2).Top;
			Microsoft.Office.Interop.Word.Shape shape5;
			object Width = (shape5 = shape2).Width;
			Microsoft.Office.Interop.Word.Shape shape6;
			object Height = (shape6 = shape2).Height;
			object Anchor = shape2.Anchor;
			Microsoft.Office.Interop.Word.Shape shape7 = shapes.AddPicture(fileName, ref Range, ref SaveWithDocument, ref LinkToFile, ref Top, ref Width, ref Height, ref Anchor);
			shape6.Height = Conversions.ToSingle(Height);
			shape5.Width = Conversions.ToSingle(Width);
			shape4.Top = Conversions.ToSingle(Top);
			shape3.Left = Conversions.ToSingle(LinkToFile);
			Microsoft.Office.Interop.Word.Shape shape8 = shape7;
			Common.A(shape8, b);
			shape8.LockAspectRatio = MsoTriState.msoTrue;
			shape8.Width = shape2.Width;
			shape2.Delete();
			shape2 = null;
			Microsoft.Office.Interop.Word.Shape shape9 = shape8;
			Anchor = RuntimeHelpers.GetObjectValue(Missing.Value);
			shape9.Select(ref Anchor);
			shape8.AlternativeText = alternativeText;
			A.Shape = shape8;
		}
		finally
		{
			shape = null;
			Microsoft.Office.Interop.Word.Shape shape8 = null;
		}
	}
}
