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

public sealed class Uncontractions : BaseTextCheck
{
	public Uncontractions()
	{
		base.RegexObj = Text.RegexUncontractions();
	}

	public override void Check(Range rng, string strText)
	{
		checked
		{
			MatchCollection matchCollection;
			if (strText.Contains(XC.A(25615)))
			{
				matchCollection = base.RegexObj.Matches(strText);
				{
					IEnumerator enumerator = matchCollection.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							Group obj = ((Match)enumerator.Current).Groups[1];
							string strFix = obj.Value.Replace(XC.A(25615), XC.A(25622));
							Range duplicate = rng.Duplicate;
							duplicate.SetRange(rng.Characters[obj.Index + 1].Start, rng.Characters[obj.Index + obj.Length].End);
							obj = null;
							Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.Uncontractions(duplicate, strFix));
							duplicate = null;
						}
						while (true)
						{
							switch (2)
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
				}
			}
			matchCollection = null;
		}
	}

	public override void Check(object shp, TextRange2 rng, string strText)
	{
		MatchCollection matchCollection;
		if (strText.Contains(XC.A(25615)))
		{
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
			matchCollection = base.RegexObj.Matches(strText);
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = matchCollection.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Group obj = ((Match)enumerator.Current).Groups[1];
					string strFix = obj.Value.Replace(XC.A(25615), XC.A(25622));
					Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.Uncontractions(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(obj.Index + 1), obj.Length), strFix));
					obj = null;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (2)
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
		matchCollection = null;
	}
}
