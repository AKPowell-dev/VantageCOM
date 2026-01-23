using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class FootnoteExplanationMissing : BaseTextError
{
	public FootnoteExplanationMissing(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strFootnote)
		: base(ErrorType.FootnoteMissing, Main.Analysis.Options.Footnotes, sld, shp, rng, blnHasFix: false)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(43414);
		((BaseError)this).Subtitle = AH.A(43471) + strFootnote + AH.A(43538);
		((BaseError)this).Tooltip = AH.A(43577);
	}
}
