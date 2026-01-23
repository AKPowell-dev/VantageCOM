using System.Collections.Generic;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class DoubleSpace : BaseTextCheck
{
	public DoubleSpace()
	{
		base.RegexObj = Text.RegexDoubleSpace();
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strText)
	{
		List<TextRange2> list = A(rng, strText, base.RegexObj, 1);
		if (list.Count > 0)
		{
			while (true)
			{
				switch (5)
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
			Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.DoubleSpace(sld, shp, list));
		}
		list = null;
	}
}
