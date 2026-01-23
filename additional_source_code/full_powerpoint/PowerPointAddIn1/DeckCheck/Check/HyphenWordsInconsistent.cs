using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class HyphenWordsInconsistent : BaseTextCheck
{
	[CompilerGenerated]
	internal sealed class CC
	{
		public string A;

		public CC(CC A)
		{
			if (A == null)
			{
				return;
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(string A)
		{
			return Operators.CompareString(A, this.A, TextCompare: false) == 0;
		}
	}

	[CompilerGenerated]
	internal sealed class DC
	{
		public string A;

		public Func<string, bool> A;

		public DC(DC A)
		{
			if (A == null)
			{
				return;
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(string A)
		{
			return Operators.CompareString(A, this.A, TextCompare: false) == 0;
		}
	}

	[CompilerGenerated]
	private new List<string> A;

	[CompilerGenerated]
	private new Dictionary<string, List<string>> A;

	[CompilerGenerated]
	private Dictionary<string, List<string>> B;

	private List<string> UniqueHyphenWords
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	private Dictionary<string, List<string>> DictHyphenWords
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	private Dictionary<string, List<string>> DictReplacementText
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public HyphenWordsInconsistent(List<string> listHyphenatedWords, List<string> listUnhyphenWords)
	{
		DC a = default(DC);
		DC CS_0024_003C_003E8__locals9 = new DC(a);
		base._002Ector();
		base.RegexObj = new Regex(Constants.REGEX_WORDS);
		UniqueHyphenWords = listHyphenatedWords.Distinct().ToList();
		DictHyphenWords = new Dictionary<string, List<string>>();
		DictReplacementText = new Dictionary<string, List<string>>();
		using List<string>.Enumerator enumerator = UniqueHyphenWords.GetEnumerator();
		CC cC = default(CC);
		while (enumerator.MoveNext())
		{
			string current = enumerator.Current;
			cC = new CC(cC);
			current = current.ToLower();
			CS_0024_003C_003E8__locals9.A = current.Replace(AH.A(17319), "");
			cC.A = current;
			List<string> list = UniqueHyphenWords.Where(cC.A).ToList();
			Func<string, bool> predicate;
			if (CS_0024_003C_003E8__locals9.A != null)
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
				predicate = CS_0024_003C_003E8__locals9.A;
			}
			else
			{
				predicate = (CS_0024_003C_003E8__locals9.A = [SpecialName] (string A) => Operators.CompareString(A, CS_0024_003C_003E8__locals9.A, TextCompare: false) == 0);
			}
			List<string> list2 = listUnhyphenWords.Where(predicate).ToList();
			List<string> list3 = new List<string>();
			List<string> list4 = new List<string>();
			string item;
			string item2;
			if (list.Count > list2.Count)
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
				list4.Add(current);
				list4.Add(CS_0024_003C_003E8__locals9.A);
				item = current + AH.A(14248) + Conversions.ToString(list.Count) + AH.A(14255);
				item2 = CS_0024_003C_003E8__locals9.A + AH.A(14248) + Conversions.ToString(list2.Count) + AH.A(14255);
			}
			else
			{
				list4.Add(CS_0024_003C_003E8__locals9.A);
				list4.Add(current);
				item = CS_0024_003C_003E8__locals9.A + AH.A(14248) + Conversions.ToString(list2.Count) + AH.A(14255);
				item2 = current + AH.A(14248) + Conversions.ToString(list.Count) + AH.A(14255);
			}
			list3.Add(item);
			list3.Add(item2);
			DictHyphenWords.Add(current, list3);
			DictReplacementText.Add(current, list4);
			list = null;
			list2 = null;
			list4 = null;
			list3 = null;
		}
		while (true)
		{
			switch (3)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = base.RegexObj.Matches(strText).GetEnumerator();
			while (enumerator.MoveNext())
			{
				Match match = (Match)enumerator.Current;
				using List<string>.Enumerator enumerator2 = UniqueHyphenWords.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					string current = enumerator2.Current;
					Group obj = match.Groups[1];
					if (Operators.CompareString(obj.Value.ToLower(), current.Replace(AH.A(17319), ""), TextCompare: false) == 0)
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
						Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.HyphenWordsInconsistent(sld, shp, para.get_Characters(checked(obj.Index + 1), obj.Length), DictHyphenWords[current], DictReplacementText[current]));
					}
					obj = null;
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_00f6;
					}
					continue;
					end_IL_00f6:
					break;
				}
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					return;
				}
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
	}
}
