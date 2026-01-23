using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class Contractions : BaseTextError
{
	public Contractions(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strFix)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).CasualWriting, sld, shp, rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.Contractions(ref val, strFix, A((List<TextRange2>)((BaseError)this).TextRanges, shp));
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		string input = ((BaseError)this).TextRanges[0].Text.Replace(AH.A(15089), "").Replace(Constants.SINGLE_QUOTE_CLOSE, "");
		string text = ((BaseError)this).ReplacementText[i];
		if (Regex.IsMatch(input, AH.A(42791)))
		{
			((BaseError)this).TextRanges[0].Text = text;
		}
		else if (Regex.IsMatch(input, AH.A(42808)))
		{
			((BaseError)this).TextRanges[0].Text = text.ToUpper();
		}
		else
		{
			((BaseError)this).TextRanges[0].Text = text.Substring(0, 1).ToUpper() + text.Substring(1, checked(text.Length - 1));
		}
	}
}
