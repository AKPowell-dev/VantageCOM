using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Check;

public sealed class GrammarEnsureInsure : BaseTextCheck
{
	public GrammarEnsureInsure()
	{
		base.RegexObj = new Regex(XC.A(24721), RegexOptions.IgnoreCase);
	}

	public override void Check(Range rng, string strText)
	{
		MatchCollection matchCollection = base.RegexObj.Matches(strText);
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = matchCollection.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Group obj = ((Match)enumerator.Current).Groups[1];
					string input = Regex.Replace(obj.Value, XC.A(24705), XC.A(24710));
					input = Regex.Replace(input, XC.A(24713), XC.A(24718));
					Range duplicate = rng.Duplicate;
					duplicate.SetRange(rng.Characters[obj.Index + 1].Start, rng.Characters[obj.Index + obj.Length].End);
					obj = null;
					Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.GrammarEnsureInsure(duplicate, input));
					duplicate = null;
				}
				while (true)
				{
					switch (3)
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
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			matchCollection = null;
		}
	}

	public override void Check(object shp, TextRange2 rng, string strText)
	{
		MatchCollection matchCollection = base.RegexObj.Matches(strText);
		IEnumerator enumerator = matchCollection.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Group obj = ((Match)enumerator.Current).Groups[1];
				string input = Regex.Replace(obj.Value, XC.A(24705), XC.A(24710));
				input = Regex.Replace(input, XC.A(24713), XC.A(24718));
				Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.GrammarEnsureInsure(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(obj.Index + 1), obj.Length), input));
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
		matchCollection = null;
	}
}
