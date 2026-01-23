using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class DoubleQuoteStyle : BaseTextCheck
{
	public DoubleQuoteStyle(DoubleSingleQuotesStyle conv)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		if ((int)conv == 0)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					base.RegexObj = new Regex(AH.A(15132));
					return;
				}
			}
		}
		base.RegexObj = new Regex(AH.A(15135) + Constants.DOUBLE_QUOTE_OPEN + Constants.DOUBLE_QUOTE_CLOSE + AH.A(15138));
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strText)
	{
		List<TextRange2> list = A(rng, strText, base.RegexObj, 0);
		if (list.Count > 0)
		{
			while (true)
			{
				switch (1)
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
			Main.Analysis.Errors.Add(new DoubleQuotesStyle(sld, shp, list, rng));
		}
		list = null;
	}
}
