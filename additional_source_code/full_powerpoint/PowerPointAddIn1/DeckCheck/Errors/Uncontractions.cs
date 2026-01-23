using System.Collections.Generic;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class Uncontractions : BaseTextError
{
	public Uncontractions(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strFix)
		: base(ErrorType.Text, (Severity)3, sld, shp, rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		BaseError val = (BaseError)(object)this;
		Errors.Uncontractions(ref val, strFix, A((List<TextRange2>)((BaseError)this).TextRanges, shp));
	}

	public override void FixAction(int i)
	{
		PowerPointAddIn1.DeckCheck.Fix.Text.ReplaceText(this, i);
	}
}
