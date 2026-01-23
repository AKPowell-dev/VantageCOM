using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing.Check;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class RepeatedWords : BaseTextCheck
{
	public RepeatedWords()
	{
		base.RegexObj = Text.RegexRepeatedWords();
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
					Match match = (Match)enumerator.Current;
					Group obj = match.Groups[1];
					string value = obj.Value;
					if (Operators.CompareString(value, XC.A(25633), TextCompare: false) != 0)
					{
						while (true)
						{
							switch (2)
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
						if (Operators.CompareString(value, XC.A(25652), TextCompare: false) != 0 && Operators.CompareString(value, XC.A(25667), TextCompare: false) != 0 && Operators.CompareString(value, XC.A(25686), TextCompare: false) != 0)
						{
							Range duplicate = rng.Duplicate;
							duplicate.SetRange(rng.Characters[obj.Index + 1].Start, rng.Characters[obj.Index + obj.Length].End);
							Main.Analysis.Errors.Add(new DuplicateWords(duplicate, match.Groups[2].Value));
							duplicate = null;
						}
					}
					obj = null;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_015a;
					}
					continue;
					end_IL_015a:
					break;
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
				Group obj = match.Groups[1];
				string value = obj.Value;
				if (Operators.CompareString(value, XC.A(25633), TextCompare: false) != 0)
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
					if (Operators.CompareString(value, XC.A(25652), TextCompare: false) != 0 && Operators.CompareString(value, XC.A(25667), TextCompare: false) != 0)
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
						if (Operators.CompareString(value, XC.A(25686), TextCompare: false) != 0)
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
							Main.Analysis.Errors.Add(new DuplicateWords(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(obj.Index + 1), obj.Length), match.Groups[2].Value));
						}
					}
				}
				obj = null;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_0127;
				}
				continue;
				end_IL_0127:
				break;
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
		matchCollection = null;
	}
}
