using System;
using System.Collections;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class GrammarIeEg : BaseTextCheck
{
	public GrammarIeEg()
	{
		base.RegexObj = new Regex(AH.A(17211) + Text.IE_EG);
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = base.RegexObj.Matches(strText).GetEnumerator();
			while (enumerator.MoveNext())
			{
				Group obj = ((Match)enumerator.Current).Groups[1];
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.GrammarIeEg(sld, shp, para.get_Characters(checked(obj.Index + 1), obj.Length)));
				obj = null;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				return;
			}
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
