using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class HyphenWordsImproper : BaseTextCheck
{
	public HyphenWordsImproper()
	{
		base.RegexObj = new Regex(AH.A(17333), RegexOptions.IgnoreCase);
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		if (!strText.Contains(AH.A(17319)))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			List<TextRange2> list = new List<TextRange2>();
			try
			{
				enumerator = base.RegexObj.Matches(strText).GetEnumerator();
				while (enumerator.MoveNext())
				{
					Group obj = ((Match)enumerator.Current).Groups[1];
					list.Add(para.get_Characters(checked(obj.Index + 1), obj.Length));
					obj = null;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_008d;
					}
					continue;
					end_IL_008d:
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			if (list.Count > 0)
			{
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.HyphenWordsImproper(sld, shp, list));
			}
			list = null;
			return;
		}
	}
}
