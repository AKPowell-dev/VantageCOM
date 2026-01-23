using System;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class Images
{
	public static void Distortion(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = shp;
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
			Main.Analysis.Errors.Add(new ImageDistortion(sld, shp, Math.Round(height / shape.Height, 2), Math.Round(width / shape.Width, 2)));
		}
		shape.Height = height;
		shape.Width = width;
		shape.LockAspectRatio = lockAspectRatio;
		shape = null;
	}

	public static void Cropping(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		try
		{
			PictureFormat pictureFormat = shp.PictureFormat;
			if (!(pictureFormat.CropTop > 0f) && !(pictureFormat.CropBottom > 0f))
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
				if (!(pictureFormat.CropLeft > 0f))
				{
					if (pictureFormat.CropRight == 0f)
					{
						goto IL_007b;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						break;
					}
				}
			}
			Main.Analysis.Errors.Add(new ImageCropping(sld, shp));
			goto IL_007b;
			IL_007b:
			pictureFormat = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	public static void AirplaneMode(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		if (PowerPointAddIn1.Shapes.AirplaneMode.IsHidden(shp))
		{
			Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.AirplaneMode(sld, shp));
		}
	}

	public static void LinkedPicture(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		if (shp.Type != MsoShapeType.msoLinkedPicture)
		{
			return;
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Main.Analysis.Errors.Add(new LinkedPicture(sld, shp));
			return;
		}
	}
}
