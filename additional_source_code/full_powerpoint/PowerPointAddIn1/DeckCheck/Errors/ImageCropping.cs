using A;
using MacabacusMacros.Proofing;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ImageCropping : BaseError
{
	public ImageCropping(Slide sld, Shape shp)
		: base(ErrorType.ImageCropping, ((Settings)Main.Analysis.Options).ImageCropping, sld, shp, blnHasFix: false)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.ImageCropping(ref val);
	}

	public override void FixAction()
	{
		Forms.WarningMessage(AH.A(33085));
	}
}
