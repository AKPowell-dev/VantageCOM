using System;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Check;

public sealed class Images
{
	public static void Distortion(Microsoft.Office.Interop.Word.Shape shp)
	{
		Microsoft.Office.Interop.Word.Shape shape = shp;
		float height = shape.Height;
		float width = shape.Width;
		MsoTriState lockAspectRatio = shape.LockAspectRatio;
		shape.LockAspectRatio = MsoTriState.msoFalse;
		shape.ScaleHeight(1f, MsoTriState.msoTrue);
		shape.ScaleWidth(1f, MsoTriState.msoTrue);
		if (Math.Round(height / shape.Height, 2) != Math.Round(width / shape.Width, 2))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Main.Analysis.Errors.Add(new ImageDistortion(shp, Math.Round(height / shape.Height, 2), Math.Round(width / shape.Width, 2)));
		}
		shape.Height = height;
		shape.Width = width;
		shape.LockAspectRatio = lockAspectRatio;
		shape = null;
	}
}
