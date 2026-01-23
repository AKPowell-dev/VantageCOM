using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class MissingOpeningCurlyBrace : BaseTextError
{
	public MissingOpeningCurlyBrace(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).PunctuationMissing, sld, shp, rng, blnHasFix: false)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(45888);
		((BaseError)this).Subtitle = A((List<TextRange2>)((BaseError)this).TextRanges, shp);
	}
}
