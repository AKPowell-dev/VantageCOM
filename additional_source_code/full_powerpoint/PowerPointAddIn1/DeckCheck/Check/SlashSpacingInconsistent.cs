using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class SlashSpacingInconsistent : BaseTextCheck
{
	public SlashSpacingInconsistent(SlashSpacing conv)
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
					base.RegexObj = new Regex(AH.A(17743));
					return;
				}
			}
		}
		base.RegexObj = new Regex(AH.A(17758));
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strText)
	{
		if (!strText.Contains(AH.A(14622)))
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			try
			{
				if (sld.Shapes.HasTitle == MsoTriState.msoTrue)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						if (sld.Shapes.Title == shp)
						{
							return;
						}
						break;
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			List<TextRange2> list = A(rng, strText, base.RegexObj, 1);
			if (list.Count > 0)
			{
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.SlashSpacingInconsistent(sld, shp, list));
			}
			list = null;
			return;
		}
	}
}
