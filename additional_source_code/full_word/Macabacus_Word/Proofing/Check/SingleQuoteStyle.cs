using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Check;

public sealed class SingleQuoteStyle : BaseTextCheck
{
	public SingleQuoteStyle(DoubleSingleQuotesStyle conv)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		if ((int)conv == 0)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					base.RegexObj = new Regex(XC.A(6379) + Constants.SINGLE_QUOTE_OPEN + Constants.SINGLE_QUOTE_CLOSE + XC.A(6382));
					base.Fix = XC.A(6376);
					return;
				}
			}
		}
		base.RegexObj = new Regex(XC.A(6376));
		base.Fix = "";
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
					Match obj = (Match)enumerator.Current;
					Range duplicate = rng.Duplicate;
					Group obj2 = obj.Groups[0];
					duplicate.SetRange(rng.Characters[obj2.Index + 1].Start, rng.Characters[obj2.Index + obj2.Length].End);
					obj2 = null;
					Main.Analysis.Errors.Add(new SingleQuotesStyle(duplicate, base.Fix));
					duplicate = null;
				}
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
					return;
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
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
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
				Group obj = ((Match)enumerator.Current).Groups[0];
				Main.Analysis.Errors.Add(new SingleQuotesStyle(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(obj.Index + 1), obj.Length), base.Fix));
				obj = null;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
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
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}
}
