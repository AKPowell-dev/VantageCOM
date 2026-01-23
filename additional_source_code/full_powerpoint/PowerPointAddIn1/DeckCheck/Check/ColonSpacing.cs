using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class ColonSpacing : BaseTextCheck
{
	public ColonSpacing(SpacesAfterColon conv)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Invalid comparison between Unknown and I4
		if ((int)conv == 1)
		{
			base.RegexObj = Text.RegexColonSingle();
			base.Fix = AH.A(15077);
		}
		else
		{
			base.RegexObj = Text.RegexColonDouble();
			base.Fix = AH.A(15084);
		}
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strText)
	{
		if (!strText.Contains(AH.A(15074)))
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
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.ColonSpacing(sld, shp, list, base.Fix));
			}
			list = null;
			return;
		}
	}
}
