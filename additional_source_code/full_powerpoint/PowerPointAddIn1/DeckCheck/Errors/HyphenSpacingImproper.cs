using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class HyphenSpacingImproper : BaseTextError
{
	public HyphenSpacingImproper(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges)
		: base(ErrorType.Text, (Severity)3, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		int count = listRanges.Count;
		string text;
		if (count == 1)
		{
			while (true)
			{
				switch (4)
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
			text = A((List<TextRange2>)((BaseError)this).TextRanges, shp);
		}
		else
		{
			text = AH.A(42825) + count + AH.A(45001);
		}
		BaseError val = (BaseError)(object)this;
		Errors.HyphenSpacingImproper(ref val, text);
	}

	public override void FixAction(int i)
	{
		PowerPointAddIn1.DeckCheck.Fix.Text.ReplaceText(this, i);
		NG.A.Application.StartNewUndoEntry();
		foreach (TextRange2 textRange in ((BaseError)this).TextRanges)
		{
			textRange.Text = ((BaseError)this).ReplacementText[i];
		}
	}
}
