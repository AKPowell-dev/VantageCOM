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

public sealed class QuestionMark : BaseTextCheck
{
	public QuestionMark()
	{
		base.RegexObj = Text.RegexQuestionMarks();
	}

	public override void Check(Range rng, string strText)
	{
		checked
		{
			MatchCollection matchCollection;
			if (strText.Contains(XC.A(2701)))
			{
				while (true)
				{
					switch (3)
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
							Match obj = (Match)enumerator.Current;
							Range duplicate = rng.Duplicate;
							Group obj2 = obj.Groups[0];
							duplicate.SetRange(rng.Characters[obj2.Index + 1].Start, rng.Characters[obj2.Index + obj2.Length].End);
							obj2 = null;
							Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.QuestionMark(duplicate, XC.A(2701)));
							duplicate = null;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_00ea;
							}
							continue;
							end_IL_00ea:
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
		if (strText.Contains(XC.A(2701)))
		{
			while (true)
			{
				switch (3)
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
					Group obj = ((Match)enumerator.Current).Groups[0];
					Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.QuestionMark(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(obj.Index + 1), obj.Length), XC.A(2701)));
					obj = null;
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_00b5;
					}
					continue;
					end_IL_00b5:
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
		}
		matchCollection = null;
	}
}
