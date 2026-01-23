using System;
using System.Collections;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;

namespace PowerPointAddIn1.DeckCheck;

public sealed class Text
{
	public static int CountCapsInTitle(string strTitle)
	{
		MatchCollection matchCollection = Regex.Matches(strTitle, AH.A(55687));
		string text = strTitle;
		IEnumerator enumerator = matchCollection.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Match match = (Match)enumerator.Current;
				text = Text.MaskText(text, match.Groups[1]);
			}
			while (true)
			{
				switch (1)
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
		text = text.Substring(1);
		matchCollection = Regex.Matches(text, AH.A(55712));
		int count = matchCollection.Count;
		matchCollection = null;
		return count;
	}
}
