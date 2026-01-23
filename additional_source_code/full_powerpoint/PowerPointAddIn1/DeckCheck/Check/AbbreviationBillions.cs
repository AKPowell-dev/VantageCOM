using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class AbbreviationBillions : BaseTextCheck
{
	public AbbreviationBillions(string convention)
	{
		if (Operators.CompareString(convention.ToLower(), AH.A(15017), TextCompare: false) != 0)
		{
			A(convention);
		}
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strText)
	{
		if (base.RegexObj == null)
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
			Regex c = Text.RegexUnits(new string[5]
			{
				AH.A(15000),
				AH.A(15005),
				AH.A(8103),
				AH.A(7905),
				AH.A(15010)
			});
			string text = A(rng, strText, c);
			if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					c = null;
					return;
				}
			}
			A(text);
			c = null;
		}
		List<TextRange2> list = A(rng, strText, base.RegexObj, 1);
		if (list.Count > 0)
		{
			Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.AbbreviationBillions(sld, shp, list, base.Fix));
		}
		list = null;
	}

	private string A(TextRange2 A, string B, Regex C)
	{
		Match match = C.Match(B);
		if (match.Success)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return match.Groups[1].Value;
				}
			}
		}
		return string.Empty;
	}

	private void A(string A)
	{
		base.RegexObj = Text.RegexBillions(A);
		base.Fix = A;
	}
}
