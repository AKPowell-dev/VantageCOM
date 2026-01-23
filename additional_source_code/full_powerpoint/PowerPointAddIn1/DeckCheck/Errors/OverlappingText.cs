using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class OverlappingText : BaseError
{
	public OverlappingText(Slide sld, Shape shp)
		: base(ErrorType.OverlappingText, Main.Analysis.Options.OverlappingText, sld, shp, blnHasFix: false)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(34383);
		((BaseError)this).Subtitle = AH.A(34422);
	}
}
