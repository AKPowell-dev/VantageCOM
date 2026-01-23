using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class FillGradient : BaseError
{
	public FillGradient(Slide sld, Shape shp)
		: base(ErrorType.FillGradient, ((Settings)Main.Analysis.Options).FillGradients, sld, shp, blnHasFix: false)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.FillGradients(ref val);
	}

	public override void FixAction(int i)
	{
	}
}
