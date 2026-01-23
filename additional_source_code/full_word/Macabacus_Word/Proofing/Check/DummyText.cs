using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using MacabacusMacros.Proofing.Check;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Check;

public sealed class DummyText : BaseTextCheck
{
	public DummyText()
	{
		base.RegexObj = Text.RegexDummyText();
	}

	public override void Check(Range para, string strText)
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
					Match obj = (Match)enumerator.Current;
					Range duplicate = para.Duplicate;
					Match match = obj;
					duplicate.SetRange(duplicate.Characters[match.Index + 1].Start, duplicate.Characters[match.Index + match.Length].End);
					match = null;
					if (duplicate.Font.Superscript == 0)
					{
						Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.DummyText(duplicate));
					}
					duplicate = null;
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
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
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = matchCollection.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Match match = (Match)enumerator.Current;
				TextRange2 rng2 = rng.get_Characters(checked(match.Index + 1), match.Length);
				if (rng.Font.Superscript == MsoTriState.msoFalse)
				{
					Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.DummyText(RuntimeHelpers.GetObjectValue(shp), rng2));
				}
				rng2 = null;
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
