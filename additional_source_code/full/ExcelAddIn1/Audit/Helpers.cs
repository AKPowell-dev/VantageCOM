using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit;

public sealed class Helpers
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<char, bool> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(char A)
		{
			return Operators.CompareString(Conversions.ToString(A), VH.A(39851), TextCompare: false) == 0;
		}
	}

	public static void MaskQuotedText(ref string strFormula)
	{
		if (!strFormula.Contains(VH.A(39830)))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			MatchCollection matchCollection = Regex.Matches(strFormula, VH.A(39833));
			enumerator = matchCollection.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					Match match = (Match)enumerator.Current;
					string source = strFormula.Substring(0, match.Index);
					Func<char, bool> predicate;
					if (_Closure_0024__.A == null)
					{
						predicate = (_Closure_0024__.A = [SpecialName] (char A) => Operators.CompareString(Conversions.ToString(A), VH.A(39851), TextCompare: false) == 0);
					}
					else
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
						predicate = _Closure_0024__.A;
					}
					if (source.Count(predicate) % 2 == 0)
					{
						strFormula = MaskFormula(match, strFormula);
					}
				}
				while (true)
				{
					switch (7)
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
				IDisposable disposable = enumerator as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
			matchCollection = null;
			return;
		}
	}

	public static void MaskSheetAndWorkbookNames(ref string strFormula, string strAddress)
	{
		if (!strAddress.Contains(VH.A(7827)))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			if (!strAddress.Contains(VH.A(39848)))
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
				string str = Strings.Split(strAddress, VH.A(7827))[0].Replace(VH.A(39851), VH.A(39854));
				try
				{
					enumerator = Regex.Matches(strFormula, Regex.Escape(str)).GetEnumerator();
					while (enumerator.MoveNext())
					{
						Match m = (Match)enumerator.Current;
						strFormula = MaskFormula(m, strFormula);
					}
					while (true)
					{
						switch (7)
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
			}
		}
	}

	public static string MaskFormula(Match m, string strFormula)
	{
		return A(strFormula, m.Index, m.Length);
	}

	public static string MaskFormula(Group g, string strFormula)
	{
		return A(strFormula, g.Index, g.Length);
	}

	private static string A(string A, int B, int C)
	{
		return Strings.Left(A, B) + new string('|', C) + Strings.Right(A, checked(A.Length - B - C));
	}

	public static List<ParenthesesPair> IdentifyParenthesesPairs(string strMaskedFormula)
	{
		List<ParenthesesPair> list = new List<ParenthesesPair>();
		Regex regex = new Regex(VH.A(39859), RegexOptions.IgnoreCase);
		checked
		{
			try
			{
				int num = 0;
				int num2 = strMaskedFormula.Length - 1;
				bool flag = false;
				int num3 = num2;
				for (int i = 0; i <= num3; i++)
				{
					string left = Conversions.ToString(strMaskedFormula[i]);
					if (Operators.CompareString(left, VH.A(39851), TextCompare: false) == 0)
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
						flag = !flag;
					}
					if (flag)
					{
						continue;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					if (Operators.CompareString(left, VH.A(39848), TextCompare: false) == 0)
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
						Match match = regex.Match(strMaskedFormula.Substring(0, i + 1));
						int intStart;
						string strFunction;
						if (match.Success)
						{
							Group obj = match.Groups[1];
							intStart = i - obj.Length;
							strFunction = obj.Value;
							obj = null;
						}
						else
						{
							intStart = i;
							strFunction = "";
						}
						match = null;
						list.Add(new ParenthesesPair(intStart, strFunction));
						num++;
					}
					else
					{
						if (Operators.CompareString(left, VH.A(39904), TextCompare: false) != 0)
						{
							continue;
						}
						List<ParenthesesPair> list2 = list;
						int num4 = list2.Count - 1;
						while (true)
						{
							if (num4 >= 0)
							{
								if (list2[num4].Length == 0)
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
									ParenthesesPair parenthesesPair = list2[num4];
									ParenthesesPair parenthesesPair2 = parenthesesPair;
									parenthesesPair2.Length = i - parenthesesPair2.StartIndex + 1;
									parenthesesPair2 = null;
									list2[num4] = parenthesesPair;
									parenthesesPair = null;
									num--;
									break;
								}
								num4 += -1;
								continue;
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								break;
							}
							break;
						}
						list2 = null;
					}
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_01cb;
					}
					continue;
					end_IL_01cb:
					break;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			regex = null;
			return list;
		}
	}
}
