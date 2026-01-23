using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class QuestionMark : BaseTextCheck
{
	public QuestionMark()
	{
		base.RegexObj = Text.RegexQuestionMarks();
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strText)
	{
		if (strText.Contains(AH.A(17524)))
		{
			List<TextRange2> list = A(rng, strText, base.RegexObj, 0);
			if (list.Count > 0)
			{
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.QuestionMark(sld, shp, list));
			}
			list = null;
		}
	}
}
