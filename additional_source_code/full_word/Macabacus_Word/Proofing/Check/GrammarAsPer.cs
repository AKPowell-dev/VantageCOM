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

public sealed class GrammarAsPer : BaseTextCheck
{
	public GrammarAsPer()
	{
		base.RegexObj = Text.RegexAsPer();
	}

	public override void Check(Range rng, string strText)
	{
		MatchCollection matchCollection = base.RegexObj.Matches(strText);
		IEnumerator enumerator = matchCollection.GetEnumerator();
		checked
		{
			try
			{
				while (enumerator.MoveNext())
				{
					Group obj = ((Match)enumerator.Current).Groups[1];
					string value = obj.Value;
					string strFix;
					if (Operators.CompareString(value, XC.A(24632), TextCompare: false) != 0)
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
						if (Operators.CompareString(value, XC.A(24645), TextCompare: false) != 0 && Operators.CompareString(value, XC.A(24658), TextCompare: false) != 0)
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
							strFix = ((Operators.CompareString(value, XC.A(24671), TextCompare: false) != 0) ? XC.A(24684) : XC.A(24698));
						}
						else
						{
							strFix = XC.A(24691);
						}
					}
					else
					{
						strFix = XC.A(24684);
					}
					Range duplicate = rng.Duplicate;
					duplicate.SetRange(rng.Characters[obj.Index + 1].Start, rng.Characters[obj.Index + obj.Length].End);
					obj = null;
					Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.GrammarAsPer(duplicate, strFix));
					duplicate = null;
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_0178;
					}
					continue;
					end_IL_0178:
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

	public override void Check(object shp, TextRange2 rng, string strText)
	{
		MatchCollection matchCollection = base.RegexObj.Matches(strText);
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			string strFix;
			Group obj;
			for (enumerator = matchCollection.GetEnumerator(); enumerator.MoveNext(); Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.GrammarAsPer(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(obj.Index + 1), obj.Length), strFix)), obj = null)
			{
				obj = ((Match)enumerator.Current).Groups[1];
				string value = obj.Value;
				if (Operators.CompareString(value, XC.A(24632), TextCompare: false) != 0)
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
					if (Operators.CompareString(value, XC.A(24645), TextCompare: false) != 0)
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
						if (Operators.CompareString(value, XC.A(24658), TextCompare: false) != 0)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								break;
							}
							strFix = ((Operators.CompareString(value, XC.A(24671), TextCompare: false) != 0) ? XC.A(24684) : XC.A(24698));
							continue;
						}
					}
					strFix = XC.A(24691);
				}
				else
				{
					strFix = XC.A(24684);
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_0145;
				}
				continue;
				end_IL_0145:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (3)
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
