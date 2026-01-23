using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class DashSpacingInconsistent : BaseTextCheck
{
	public DashSpacingInconsistent(DashSpacing conv)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Invalid comparison between Unknown and I4
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
					base.RegexObj = new Regex(AH.A(15095));
					return;
				}
			}
		}
		base.RegexObj = new Regex(AH.A(15110));
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strText)
	{
		if (!strText.Contains(AH.A(15092)))
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			List<TextRange2> list = A(rng, strText, base.RegexObj, 1);
			if (list.Count > 0)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					break;
				}
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.DashSpacingInconsistent(sld, shp, list));
			}
			list = null;
			return;
		}
	}
}
