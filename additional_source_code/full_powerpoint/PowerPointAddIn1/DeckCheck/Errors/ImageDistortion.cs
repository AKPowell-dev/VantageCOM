using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ImageDistortion : BaseError
{
	public ImageDistortion(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, double dblHeight, double dblWidth)
		: base(ErrorType.ImageDistortion, ((Settings)Main.Analysis.Options).ImageDistortion, sld, shp, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.ImageDistortion(ref val, dblHeight, dblWidth);
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		Microsoft.Office.Interop.PowerPoint.Shape shape = base.Shape;
		MsoTriState lockAspectRatio = shape.LockAspectRatio;
		shape.LockAspectRatio = MsoTriState.msoFalse;
		float factor;
		if (i == 0)
		{
			float width = shape.Width;
			shape.ScaleWidth(1f, MsoTriState.msoTrue);
			factor = width / shape.Width;
		}
		else
		{
			float height = shape.Height;
			shape.ScaleHeight(1f, MsoTriState.msoTrue);
			factor = height / shape.Height;
		}
		shape.ScaleHeight(factor, MsoTriState.msoTrue);
		shape.ScaleWidth(factor, MsoTriState.msoTrue);
		shape.LockAspectRatio = lockAspectRatio;
		shape = null;
	}
}
