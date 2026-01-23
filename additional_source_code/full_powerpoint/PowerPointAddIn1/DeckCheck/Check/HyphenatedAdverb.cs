using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class HyphenatedAdverb : BaseTextCheck
{
	public HyphenatedAdverb()
	{
		base.RegexObj = Text.RegexHyphenAdverb();
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strText)
	{
		if (!strText.Contains(AH.A(17319)))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			List<TextRange2> list = new List<TextRange2>();
			enumerator = base.RegexObj.Matches(strText).GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					Group obj = ((Match)enumerator.Current).Groups[1];
					list.Add(rng.get_Characters(checked(obj.Index + 1), obj.Length));
					obj = null;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_008f;
					}
					continue;
					end_IL_008f:
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
					switch (4)
					{
					case 0:
						continue;
					}
					break;
				}
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.HyphenatedAdverb(sld, shp, list));
			}
			list = null;
			return;
		}
	}
}
