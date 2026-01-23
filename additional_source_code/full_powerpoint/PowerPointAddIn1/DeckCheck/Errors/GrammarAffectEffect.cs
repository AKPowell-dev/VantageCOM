using System.Collections.Generic;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class GrammarAffectEffect : BaseTextError
{
	public GrammarAffectEffect(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strFix)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).ConfusedWords, sld, shp, rng, blnHasFix: true)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.GrammarAffectEffect(ref val, strFix, A((List<TextRange2>)((BaseError)this).TextRanges, shp));
	}

	public override void FixAction(int i)
	{
		PowerPointAddIn1.DeckCheck.Fix.Text.ReplaceText(this, i);
	}
}
