using System;
using System.Collections;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class Uncontractions : BaseTextCheck
{
	public Uncontractions()
	{
		base.RegexObj = Text.RegexUncontractions();
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		if (!strText.Contains(AH.A(17819)))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			try
			{
				enumerator = base.RegexObj.Matches(strText).GetEnumerator();
				while (enumerator.MoveNext())
				{
					Group obj = ((Match)enumerator.Current).Groups[1];
					string strFix = obj.Value.Replace(AH.A(17819), AH.A(17826));
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.Uncontractions(sld, shp, para.get_Characters(checked(obj.Index + 1), obj.Length), strFix));
					obj = null;
				}
				return;
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
		}
	}
}
