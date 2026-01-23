using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using MacabacusMacros.Proofing.CorporateDictionary;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class CorporateDictionary : BaseTextCheck
{
	[CompilerGenerated]
	private new List<Rule> A;

	private List<Rule> Rules
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
	}

	public CorporateDictionary(List<Rule> _rules)
	{
		A = _rules;
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		List<TextRange2> list;
		using (List<Rule>.Enumerator enumerator = Rules.GetEnumerator())
		{
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				Rule current = enumerator.Current;
				list = new List<TextRange2>();
				try
				{
					enumerator2 = current.SearchRegex.Matches(strText).GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Match match = (Match)enumerator2.Current;
						if (current.ReplaceWith.Contains(match.Groups[1].Value))
						{
							continue;
						}
						while (true)
						{
							switch (7)
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
						Group obj = match.Groups[1];
						list.Add(para.get_Characters(checked(obj.Index + 1), obj.Length));
						obj = null;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
				if (list.Count > 0)
				{
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.CorporateDictionary(sld, shp, list, current));
				}
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_011e;
				}
				continue;
				end_IL_011e:
				break;
			}
		}
		list = null;
	}
}
