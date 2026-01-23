using System;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class MissingStraightQuotes : BaseTextCheck
{
	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		if (!strText.Contains(AH.A(15132)))
		{
			return;
		}
		try
		{
			if (Strings.Split(strText, AH.A(15132)).Length % 2 == 0)
			{
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.MissingStraightQuotes(sld, shp, para.get_Characters(checked(strText.LastIndexOf('"') + 1), 1)));
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
