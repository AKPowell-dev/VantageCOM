using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
using A;
using MacabacusMacros.ImportExport;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.Library2.Insert;

public sealed class Images
{
	internal struct FD
	{
		public float A;

		public float B;

		public float C;

		public float D;

		public int A;
	}

	internal static Microsoft.Office.Interop.PowerPoint.Shape A(Slide A, string B)
	{
		return A.Shapes.AddPicture(B, MsoTriState.msoFalse, MsoTriState.msoTrue, 0f, 0f);
	}

	internal static bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		if (A.Type == MsoShapeType.msoPlaceholder)
		{
			return ExcelToPowerPoint.IsPictureHolder(A.PlaceholderFormat.Type);
		}
		return false;
	}

	internal static Microsoft.Office.Interop.PowerPoint.Shape A(Microsoft.Office.Interop.PowerPoint.Application A, Slide B, string C)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = Images.A(B, C);
		if (shape.Type == MsoShapeType.msoPlaceholder)
		{
			while (true)
			{
				switch (7)
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
			shape.Cut();
			A.ActiveWindow.View.Paste();
			shape = A.ActiveWindow.Selection.ShapeRange[1];
			Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
			PictureFormat pictureFormat = shape2.PictureFormat;
			pictureFormat.CropBottom = 0f;
			pictureFormat.CropLeft = 0f;
			pictureFormat.CropRight = 0f;
			pictureFormat.CropTop = 0f;
			_ = null;
			shape2.ScaleHeight(1f, MsoTriState.msoTrue);
			shape2.ScaleWidth(1f, MsoTriState.msoTrue);
			shape2.Top = 0f;
			shape2.Left = 0f;
			_ = null;
		}
		return shape;
	}

	internal static void A(Microsoft.Office.Interop.PowerPoint.Shape A, int B)
	{
		while (A.ZOrderPosition > B)
		{
			A.ZOrder(MsoZOrderCmd.msoSendBackward);
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
			return;
		}
	}

	internal static void A(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Application B)
	{
		A.Select();
		B.CommandBars.ExecuteMso(AH.A(58871));
		System.Windows.Forms.Application.DoEvents();
		A.Select();
	}

	internal static Microsoft.Office.Interop.PowerPoint.Shape A(Microsoft.Office.Interop.PowerPoint.Shape A, Slide B)
	{
		if (A.PlaceholderFormat.ContainedType != MsoShapeType.msoPicture)
		{
			if (A.PlaceholderFormat.ContainedType != (MsoShapeType)PowerPointAddIn1.Shapes.Images.A)
			{
				goto IL_007a;
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
		}
		int zOrderPosition = A.ZOrderPosition;
		A.Delete();
		A = B.Shapes[B.Shapes.Count];
		Images.A(A, zOrderPosition);
		A.Select();
		goto IL_007a;
		IL_007a:
		return A;
	}

	internal static void A(Slide A, ref int B, ref Dictionary<int, FD> C)
	{
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.Shapes.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
					if (!Images.A(shape))
					{
						continue;
					}
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
					if (shape.PlaceholderFormat.ContainedType != MsoShapeType.msoAutoShape)
					{
						continue;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						break;
					}
					B++;
					C.Add(shape.ZOrderPosition, new FD
					{
						A = shape.Top,
						B = shape.Left,
						C = shape.Height,
						D = shape.Width,
						A = shape.ZOrderPosition
					});
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
		}
	}

	internal static void A(Slide A, Dictionary<Microsoft.Office.Interop.PowerPoint.Shape, FD> B)
	{
		foreach (KeyValuePair<Microsoft.Office.Interop.PowerPoint.Shape, FD> item in B)
		{
			item.Key.PickUp();
			item.Key.Delete();
			Microsoft.Office.Interop.PowerPoint.Shape shape = A.Shapes[A.Shapes.Count];
			Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
			shape2.Top = item.Value.A;
			shape2.Left = item.Value.B;
			shape2.Height = item.Value.C;
			shape2.Width = item.Value.D;
			Images.A(shape, item.Value.A);
			shape2.Apply();
			_ = null;
			shape = null;
		}
	}
}
