using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class SentenceSpacing : BaseTextCheck
{
	public SentenceSpacing(SpacesBetweenSentences conv)
	{
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0020: Invalid comparison between Unknown and I4
		string text = AH.A(17591);
		string text2 = AH.A(17684);
		if ((int)conv == 1)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					base.RegexObj = new Regex(text + AH.A(17707) + text2);
					base.Fix = AH.A(17718);
					return;
				}
			}
		}
		base.RegexObj = new Regex(text + AH.A(17725) + text2);
		base.Fix = AH.A(17738);
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strText)
	{
		if (!strText.Contains(AH.A(14417)))
		{
			return;
		}
		List<TextRange2> list = A(rng, strText, base.RegexObj, 2);
		if (list.Count > 0)
		{
			while (true)
			{
				switch (3)
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
			Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.SentenceSpacing(sld, shp, list, base.Fix));
		}
		list = null;
	}
}
