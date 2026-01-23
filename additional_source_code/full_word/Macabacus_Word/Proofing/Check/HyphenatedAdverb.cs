using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing.Check;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Check;

public sealed class HyphenatedAdverb : BaseTextCheck
{
	public HyphenatedAdverb()
	{
		base.RegexObj = Text.RegexHyphenAdverb();
	}

	public override void Check(Range rng, string strText)
	{
		checked
		{
			MatchCollection matchCollection;
			if (strText.Contains(XC.A(24911)))
			{
				while (true)
				{
					switch (5)
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
				matchCollection = base.RegexObj.Matches(strText);
				foreach (Match item in matchCollection)
				{
					Range duplicate = rng.Duplicate;
					Group obj2 = item.Groups[1];
					duplicate.SetRange(rng.Characters[obj2.Index + 1].Start, rng.Characters[obj2.Index + obj2.Length].End);
					obj2 = null;
					Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.HyphenatedAdverb(duplicate, XC.A(24918)));
					duplicate = null;
				}
			}
			matchCollection = null;
		}
	}

	public override void Check(object shp, TextRange2 rng, string strText)
	{
		MatchCollection matchCollection;
		if (strText.Contains(XC.A(24911)))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			matchCollection = base.RegexObj.Matches(strText);
			{
				IEnumerator enumerator = matchCollection.GetEnumerator();
				try
				{
					while (enumerator.MoveNext())
					{
						Group obj = ((Match)enumerator.Current).Groups[1];
						Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.HyphenatedAdverb(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(obj.Index + 1), obj.Length), XC.A(24918)));
						obj = null;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_00af;
						}
						continue;
						end_IL_00af:
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
			}
		}
		matchCollection = null;
	}
}
