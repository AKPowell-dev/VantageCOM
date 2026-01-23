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

public sealed class GrammarMyriadOf : BaseTextCheck
{
	public GrammarMyriadOf()
	{
		base.RegexObj = Text.RegexMyriadOf();
	}

	public override void Check(Range rng, string strText)
	{
		MatchCollection matchCollection = base.RegexObj.Matches(strText);
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				string strFix;
				Group obj;
				Range duplicate;
				for (enumerator = matchCollection.GetEnumerator(); enumerator.MoveNext(); duplicate = rng.Duplicate, duplicate.SetRange(rng.Characters[obj.Index + 1].Start, rng.Characters[obj.Index + obj.Length].End), obj = null, Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.GrammarMyriadOf(duplicate, strFix)), duplicate = null)
				{
					obj = ((Match)enumerator.Current).Groups[1];
					string value = obj.Value;
					if (Operators.CompareString(value, XC.A(24796), TextCompare: false) != 0)
					{
						if (Operators.CompareString(value, XC.A(24815), TextCompare: false) != 0)
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
							if (Operators.CompareString(value, XC.A(24834), TextCompare: false) != 0)
							{
								if (Operators.CompareString(value, XC.A(24853), TextCompare: false) != 0)
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
									strFix = XC.A(24872);
								}
								else
								{
									strFix = XC.A(24898);
								}
								continue;
							}
						}
						strFix = XC.A(24885);
					}
					else
					{
						strFix = XC.A(24872);
					}
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_0176;
					}
					continue;
					end_IL_0176:
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

	public override void Check(object shp, TextRange2 rng, string strText)
	{
		MatchCollection matchCollection = base.RegexObj.Matches(strText);
		IEnumerator enumerator = matchCollection.GetEnumerator();
		try
		{
			string strFix;
			Group obj;
			for (; enumerator.MoveNext(); Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.GrammarMyriadOf(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(obj.Index + 1), obj.Length), strFix)), obj = null)
			{
				obj = ((Match)enumerator.Current).Groups[1];
				string value = obj.Value;
				if (Operators.CompareString(value, XC.A(24796), TextCompare: false) != 0)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (Operators.CompareString(value, XC.A(24815), TextCompare: false) != 0)
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
						if (Operators.CompareString(value, XC.A(24834), TextCompare: false) != 0)
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
							strFix = ((Operators.CompareString(value, XC.A(24853), TextCompare: false) != 0) ? XC.A(24872) : XC.A(24898));
							continue;
						}
					}
					strFix = XC.A(24885);
				}
				else
				{
					strFix = XC.A(24872);
				}
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_014c;
				}
				continue;
				end_IL_014c:
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
