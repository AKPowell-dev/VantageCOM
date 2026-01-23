using MacabacusMacros;
using MacabacusMacros.Libraries.Versioning;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Library2.Versioning.Replace;

public sealed class Common
{
	internal struct Q
	{
		public WdWrapType A;

		public WdWrapSideType A;

		public float A;

		public float B;

		public float C;

		public float D;

		public bool A;

		public WdRelativeVerticalPosition A;

		public WdRelativeHorizontalPosition A;

		public float E;

		public float F;

		public float G;

		public float H;

		public float I;

		public int A;
	}

	internal static string A(ContentItem A)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		return CloudStorage.FillPlaceholdersInPath(A.ContentInfo.ContentPath);
	}

	internal static Q A(Microsoft.Office.Interop.Word.Shape A)
	{
		WrapFormat wrapFormat = A.WrapFormat;
		return new Q
		{
			A = wrapFormat.Type,
			A = wrapFormat.Side,
			A = wrapFormat.DistanceTop,
			B = wrapFormat.DistanceBottom,
			C = wrapFormat.DistanceLeft,
			D = wrapFormat.DistanceRight,
			A = (wrapFormat.AllowOverlap != 0),
			A = A.RelativeVerticalPosition,
			A = A.RelativeHorizontalPosition,
			E = A.Rotation,
			F = A.Height,
			G = A.Width,
			H = A.Top,
			I = A.Left
		};
	}

	internal static void A(Microsoft.Office.Interop.Word.Shape A, Q B)
	{
		if (B.A == WdWrapType.wdWrapInline)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					A.WrapFormat.Type = B.A;
					A.Height = B.F;
					A.Width = B.G;
					_ = null;
					return;
				}
			}
		}
		Microsoft.Office.Interop.Word.Shape shape = A;
		WrapFormat wrapFormat = shape.WrapFormat;
		wrapFormat.Type = B.A;
		wrapFormat.Side = B.A;
		wrapFormat.DistanceTop = B.A;
		wrapFormat.DistanceBottom = B.B;
		wrapFormat.DistanceLeft = B.C;
		wrapFormat.DistanceRight = B.D;
		wrapFormat.AllowOverlap = 0 - (B.A ? 1 : 0);
		_ = null;
		shape.RelativeVerticalPosition = B.A;
		shape.RelativeHorizontalPosition = B.A;
		shape.Rotation = B.E;
		shape.Height = B.F;
		shape.Width = B.G;
		shape.Top = B.H;
		shape.Left = B.I;
		int zOrderPosition = shape.ZOrderPosition;
		while (true)
		{
			if (shape.ZOrderPosition > B.A)
			{
				shape.ZOrder(MsoZOrderCmd.msoSendBackward);
				if (shape.ZOrderPosition == zOrderPosition)
				{
					break;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_011c;
					}
					continue;
					end_IL_011c:
					break;
				}
				zOrderPosition = shape.ZOrderPosition;
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
			break;
		}
		shape = null;
	}
}
