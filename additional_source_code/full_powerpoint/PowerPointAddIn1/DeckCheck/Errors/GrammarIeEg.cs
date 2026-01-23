using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class GrammarIeEg : BaseTextError
{
	public GrammarIeEg(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).GrammarIeEg, sld, shp, rng, blnHasFix: true)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.GrammarIeEg(ref val, A((List<TextRange2>)((BaseError)this).TextRanges, shp));
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		TextRange2 textRange = ((BaseError)this).TextRanges[0];
		if (textRange.Text.Contains(AH.A(44671)))
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			textRange.Text = textRange.Text.Replace(AH.A(44671), AH.A(44680));
		}
		else
		{
			textRange.Text = textRange.Text.Replace(AH.A(44680), AH.A(44671));
		}
		textRange = null;
	}
}
