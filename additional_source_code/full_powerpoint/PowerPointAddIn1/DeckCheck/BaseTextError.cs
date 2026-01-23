using System.Collections.Generic;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck;

public class BaseTextError : BaseError
{
	public BaseTextError(ErrorType errType, Severity sev, Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRange, bool blnHasFix, bool blnCanFixMultiple = false)
		: base(errType, sev, sld, shp, blnHasFix, blnCanFixMultiple)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).TextRanges = listRange;
	}

	public BaseTextError(ErrorType errType, Severity sev, Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, bool blnHasFix, bool blnCanFixMultiple = false)
		: base(errType, sev, sld, shp, blnHasFix, blnCanFixMultiple)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).TextRanges = new List<TextRange2>(new TextRange2[1] { rng });
	}
}
