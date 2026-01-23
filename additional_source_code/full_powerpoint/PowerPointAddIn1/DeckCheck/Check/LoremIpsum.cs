using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class LoremIpsum : BaseTextCheck
{
	public LoremIpsum()
	{
		base.RegexObj = Text.RegexLoremIpsum();
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		List<TextRange2> list = new List<TextRange2>();
		IEnumerator enumerator = base.RegexObj.Matches(strText).GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Match match = (Match)enumerator.Current;
				list.Add(para.get_Characters(checked(match.Index + 1), match.Length));
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
				break;
			}
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
		if (list.Count > 0)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.LoremIpsum(sld, shp, list));
		}
		list = null;
	}
}
