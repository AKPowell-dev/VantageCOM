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

public sealed class UnnecessaryPeriods : BaseTextCheck
{
	public UnnecessaryPeriods()
	{
		base.RegexObj = Text.RegexUnnecessaryPeriods();
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strText)
	{
		if (!strText.Contains(AH.A(14417)))
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
					TextRange2 textRange = rng.get_Characters(checked(obj.Index + 1), obj.Length);
					obj = null;
					if (textRange.Text.Contains(AH.A(14417)))
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
						list.Add(textRange);
					}
					textRange = null;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (5)
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
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					break;
				}
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.UnnecessaryPeriods(sld, shp, list));
			}
			list = null;
			return;
		}
	}
}
