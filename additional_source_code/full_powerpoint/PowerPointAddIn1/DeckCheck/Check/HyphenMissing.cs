using System.Collections.Generic;
using System.Text.RegularExpressions;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class HyphenMissing : BaseTextCheck
{
	public HyphenMissing()
	{
		base.RegexObj = Text.RegexHyphenMissing();
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strText)
	{
		List<TextRange2> list = new List<TextRange2>();
		foreach (Match item in base.RegexObj.Matches(strText))
		{
			Group obj = item.Groups[1];
			list.Add(rng.get_Characters(checked(obj.Index + 1), obj.Length));
			obj = null;
		}
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
			Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.HyphenMissing(sld, shp, list));
		}
		list = null;
	}
}
