using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class SlideTitleMissing : BaseError
{
	public SlideTitleMissing(Slide sld, Shape shp)
		: base(ErrorType.TitleMissing, Main.Analysis.Options.SlideTitles, sld, shp, blnHasFix: false)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(37248);
		((BaseError)this).Subtitle = AH.A(37287);
	}
}
