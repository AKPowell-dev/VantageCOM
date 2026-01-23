using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Macabacus_Word.Values;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.TextOps.Redaction.Redactors;

public sealed class PictureRedactor
{
	private static bool m_A = false;

	public static void RedactInlinePicture(InlineShape shp)
	{
		InlineShapeValue inlineShapeValue = new InlineShapeValue(shp);
		shp.Delete();
		InlineShapes inlineShapes = inlineShapeValue.Range.InlineShapes;
		string fileName = A();
		object LinkToFile = RuntimeHelpers.GetObjectValue(Missing.Value);
		object SaveWithDocument = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Range = RuntimeHelpers.GetObjectValue(Missing.Value);
		InlineShape inlineShape = inlineShapes.AddPicture(fileName, ref LinkToFile, ref SaveWithDocument, ref Range);
		inlineShape.Height = inlineShapeValue.Height;
		inlineShape.Width = inlineShapeValue.Width;
	}

	private static void A(Microsoft.Office.Interop.Word.Shape A)
	{
		A.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
	}

	public static void RedactFloatingPicture(Microsoft.Office.Interop.Word.Shape shp)
	{
		A(shp);
		ShapeValue shapeValue = new ShapeValue(shp);
		shp.Delete();
		Microsoft.Office.Interop.Word.Shapes shapes = PC.A.Application.ActiveDocument.Shapes;
		string fileName = A();
		object LinkToFile = RuntimeHelpers.GetObjectValue(Missing.Value);
		object SaveWithDocument = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Left = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Top = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Width = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Height = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Anchor = shapeValue.Anchor;
		Microsoft.Office.Interop.Word.Shape shape = shapes.AddPicture(fileName, ref LinkToFile, ref SaveWithDocument, ref Left, ref Top, ref Width, ref Height, ref Anchor);
		A(shape, shapeValue);
		int zOrderPosition = shape.ZOrderPosition;
		while (true)
		{
			if (shape.ZOrderPosition > shapeValue.ZOrderPosition)
			{
				shape.ZOrder(MsoZOrderCmd.msoSendBackward);
				if (zOrderPosition == shape.ZOrderPosition)
				{
					break;
				}
				zOrderPosition = shape.ZOrderPosition;
				continue;
			}
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
			break;
		}
		shapeValue = null;
		shape = null;
	}

	private static void A(Microsoft.Office.Interop.Word.Shape A, ShapeValue B)
	{
		WrapFormat wrapFormat = A.WrapFormat;
		wrapFormat.Type = B.WrapType;
		wrapFormat.Side = B.WrapSide;
		wrapFormat.DistanceTop = B.DistanceTop;
		wrapFormat.DistanceBottom = B.DistanceBottom;
		wrapFormat.DistanceLeft = B.DistanceLeft;
		wrapFormat.DistanceRight = B.DistanceRight;
		wrapFormat.AllowOverlap = 0 - (B.AllowOverlap ? 1 : 0);
		_ = null;
		A.RelativeVerticalPosition = B.RelativeVerticalPosition;
		A.RelativeHorizontalPosition = B.RelativeHorizontalPosition;
		A.Rotation = B.Rotation;
		A.Height = B.Height;
		A.Width = B.Width;
		A.Top = B.Top;
		A.Left = B.Left;
		_ = null;
	}

	private static void A(string A, int B, int C)
	{
		Bitmap bitmap = new Bitmap(B, C);
		try
		{
			using (Graphics graphics = Graphics.FromImage(bitmap))
			{
				graphics.Clear(Color.Black);
			}
			bitmap.Save(A, ImageFormat.Png);
		}
		finally
		{
			if (bitmap != null)
			{
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
					((IDisposable)bitmap).Dispose();
					break;
				}
			}
		}
	}

	private static string A()
	{
		string text = Path.GetTempPath() + XC.A(19643);
		if (!PictureRedactor.m_A)
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
			A(text, 1000, 1000);
			PictureRedactor.m_A = true;
		}
		return text;
	}
}
