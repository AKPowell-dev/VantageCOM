using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class SlideTitleCapitalization : BaseTextError
{
	public SlideTitleCapitalization(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strFix)
		: base(ErrorType.Text, Main.Analysis.Options.SlideTitles, sld, shp, rng, blnHasFix: true)
	{
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).DisplayText = ((BaseError)this).ConvertFixStringToFixes(strFix);
		((BaseError)this).ReplacementText = new List<string>(new string[1] { strFix });
		((BaseError)this).Title = AH.A(36761);
		((BaseError)this).Subtitle = AH.A(36828);
		((BaseError)this).Tooltip = AH.A(37034);
	}

	public override void FixAction(int i)
	{
		PowerPointAddIn1.DeckCheck.Fix.Text.ReplaceText(this, i);
	}
}
