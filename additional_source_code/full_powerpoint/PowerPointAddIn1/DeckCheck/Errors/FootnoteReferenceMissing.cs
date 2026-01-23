using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class FootnoteReferenceMissing : BaseError
{
	public FootnoteReferenceMissing(Slide sld, Shape shp, int intFootnote)
		: base(ErrorType.FootnotesSequence, Main.Analysis.Options.Footnotes, sld, shp, blnHasFix: false)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(43821);
		((BaseError)this).Subtitle = string.Format(AH.A(43874), intFootnote);
		((BaseError)this).Tooltip = AH.A(44004);
	}
}
