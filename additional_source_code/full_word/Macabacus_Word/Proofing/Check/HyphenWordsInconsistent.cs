using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class HyphenWordsInconsistent : BaseTextCheck
{
	[CompilerGenerated]
	internal sealed class EC
	{
		public string A;

		public EC(EC A)
		{
			if (A == null)
			{
				return;
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
	internal sealed class FC
	{
		public string A;

		public Func<string, bool> A;

		public FC(FC A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal bool A(string A)
		{
			return Operators.CompareString(A, this.A, TextCompare: false) == 0;
		}
	}

	private List<string> m_A;

	private Dictionary<string, List<string>> m_A;

	private Dictionary<string, List<string>> m_B;

	private List<string> A
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	private Dictionary<string, List<string>> A
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	private Dictionary<string, List<string>> B
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
		}
	}

	public HyphenWordsInconsistent(List<string> listHyphenatedWords, List<string> listUnhyphenWords)
	{
		FC a = default(FC);
		FC CS_0024_003C_003E8__locals9 = new FC(a);
		base._002Ector();
		base.RegexObj = new Regex(Constants.REGEX_WORDS);
		this.A = listHyphenatedWords.Distinct().ToList();
		A = new Dictionary<string, List<string>>();
		B = new Dictionary<string, List<string>>();
		using List<string>.Enumerator enumerator = this.A.GetEnumerator();
		EC eC = default(EC);
		while (enumerator.MoveNext())
		{
			string current = enumerator.Current;
			eC = new EC(eC);
			current = current.ToLower();
			CS_0024_003C_003E8__locals9.A = current.Replace(XC.A(6388), "");
			eC.A = current;
			List<string> list = this.A.Where(eC.A).ToList();
			Func<string, bool> predicate;
			if (CS_0024_003C_003E8__locals9.A != null)
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
				list4.Add(current);
				list4.Add(CS_0024_003C_003E8__locals9.A);
				item = current + XC.A(22691) + Conversions.ToString(list.Count) + XC.A(20696);
				item2 = CS_0024_003C_003E8__locals9.A + XC.A(22691) + Conversions.ToString(list2.Count) + XC.A(20696);
			}
			else
			{
				list4.Add(CS_0024_003C_003E8__locals9.A);
				list4.Add(current);
				item = CS_0024_003C_003E8__locals9.A + XC.A(22691) + Conversions.ToString(list2.Count) + XC.A(20696);
				item2 = current + XC.A(22691) + Conversions.ToString(list.Count) + XC.A(20696);
			}
			list3.Add(item);
			list3.Add(item2);
			A.Add(current, list3);
			B.Add(current, list4);
			list = null;
			list2 = null;
			list4 = null;
			list3 = null;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				break;
			default:
				return;
			}
		}
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
					using List<string>.Enumerator enumerator2 = this.A.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						string current = enumerator2.Current;
						Group obj = match.Groups[1];
						if (Operators.CompareString(obj.Value.ToLower(), current.Replace(XC.A(6388), ""), TextCompare: false) == 0)
						{
							Range duplicate = rng.Duplicate;
							duplicate.SetRange(rng.Characters[obj.Index + 1].Start, rng.Characters[obj.Index + obj.Length].End);
							Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.HyphenWordsInconsistent(duplicate, A[current], B[current]));
							duplicate = null;
						}
						obj = null;
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
				using List<string>.Enumerator enumerator2 = this.A.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					string current = enumerator2.Current;
					Group obj = match.Groups[1];
					if (Operators.CompareString(obj.Value.ToLower(), current.Replace(XC.A(6388), ""), TextCompare: false) == 0)
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
						Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.HyphenWordsInconsistent(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(obj.Index + 1), obj.Length), A[current], B[current]));
					}
					obj = null;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_0103;
					}
					continue;
					end_IL_0103:
					break;
				}
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
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		matchCollection = null;
	}
}
