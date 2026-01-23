using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class MultipleFontFamilies
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<string, C<string>> A;

		public static Func<string, C<string>> B;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal C<string> A(string A)
		{
			return new C<string>(A.ToString());
		}

		[SpecialName]
		internal C<string> B(string A)
		{
			return new C<string>(A.ToString());
		}
	}

	public void Check(Range rng)
	{
		try
		{
			if (!string.IsNullOrEmpty(rng.Font.Name))
			{
				return;
			}
			List<string> list = new List<string>();
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = rng.Words.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					if (Operators.CompareString(Conversions.ToString(range.Text.Last()), XC.A(18458), TextCompare: false) == 0)
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
						range.SetRange(range.Start, checked(range.End - 1));
					}
					if (!Regex.IsMatch(range.Text, XC.A(22714)))
					{
						continue;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						break;
					}
					if (!string.IsNullOrEmpty(range.Font.Name))
					{
						list.Add(range.Font.Name);
					}
				}
				while (true)
				{
					switch (2)
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
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			List<string> list2 = new List<string>();
			List<string> list3 = new List<string>();
			IEnumerator<IGrouping<C<string>, string>> enumerator2 = default(IEnumerator<IGrouping<C<string>, string>>);
			try
			{
				List<string> source = list;
				Func<string, C<string>> keySelector;
				if (_Closure_0024__.A == null)
				{
					keySelector = (_Closure_0024__.A = [SpecialName] (string A) => new C<string>(A.ToString()));
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
					keySelector = _Closure_0024__.A;
				}
				enumerator2 = source.GroupBy(keySelector).GetEnumerator();
				while (enumerator2.MoveNext())
				{
					IGrouping<C<string>, string> current = enumerator2.Current;
					list2.Add(current.Key.Font);
					list3.Add(current.Key.Font + XC.A(22691) + current.Count() + XC.A(20696));
					current = null;
				}
			}
			finally
			{
				if (enumerator2 != null)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						enumerator2.Dispose();
						break;
					}
				}
			}
			Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.MultipleFontFamilies(rng.Duplicate, list3, string.Join(XC.A(22698), list3.ToArray()), list2));
			list = null;
			list2 = null;
			list3 = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public void Check(object shp, TextRange2 rng)
	{
		try
		{
			if (!string.IsNullOrEmpty(rng.Font.Name))
			{
				return;
			}
			IEnumerator enumerator = default(IEnumerator);
			IEnumerator<IGrouping<C<string>, string>> enumerator2 = default(IEnumerator<IGrouping<C<string>, string>>);
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
				List<string> list = new List<string>();
				try
				{
					enumerator = rng.get_Words(-1, -1).GetEnumerator();
					while (enumerator.MoveNext())
					{
						TextRange2 textRange = (TextRange2)enumerator.Current;
						if (Operators.CompareString(Conversions.ToString(textRange.Text.Last()), XC.A(18458), TextCompare: false) == 0)
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
							textRange = checked(rng.get_Characters(textRange.Start - rng.Start + 1, textRange.Length - 1));
						}
						if (!Regex.IsMatch(textRange.Text, XC.A(22714)))
						{
							continue;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							break;
						}
						if (string.IsNullOrEmpty(textRange.Font.Name))
						{
							continue;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							break;
						}
						list.Add(textRange.Font.Name);
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
				List<string> list2 = new List<string>();
				List<string> list3 = new List<string>();
				try
				{
					enumerator2 = (from A in list
						group A by new C<string>(A.ToString())).GetEnumerator();
					while (enumerator2.MoveNext())
					{
						IGrouping<C<string>, string> current = enumerator2.Current;
						list2.Add(current.Key.Font);
						list3.Add(current.Key.Font + XC.A(22691) + current.Count() + XC.A(20696));
						current = null;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_01f5;
						}
						continue;
						end_IL_01f5:
						break;
					}
				}
				finally
				{
					if (enumerator2 != null)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							enumerator2.Dispose();
							break;
						}
					}
				}
				Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.MultipleFontFamilies(RuntimeHelpers.GetObjectValue(shp), rng, list3, string.Join(XC.A(22698), list3.ToArray()), list2));
				list = null;
				list2 = null;
				list3 = null;
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
