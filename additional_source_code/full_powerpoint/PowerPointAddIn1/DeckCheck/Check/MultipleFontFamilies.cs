using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class MultipleFontFamilies
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<string, S<string>> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal S<string> A(string A)
		{
			return new S<string>(A.ToString());
		}
	}

	public void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para)
	{
		try
		{
			if (!string.IsNullOrEmpty(para.Font.Name))
			{
				return;
			}
			IEnumerator enumerator = default(IEnumerator);
			IEnumerator<IGrouping<S<string>, string>> enumerator2 = default(IEnumerator<IGrouping<S<string>, string>>);
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
				List<string> list = new List<string>();
				try
				{
					enumerator = para.get_Words(-1, -1).GetEnumerator();
					while (enumerator.MoveNext())
					{
						TextRange2 textRange = (TextRange2)enumerator.Current;
						if (Operators.CompareString(Conversions.ToString(textRange.Text.Last()), AH.A(14625), TextCompare: false) == 0)
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
							textRange = checked(para.get_Characters(textRange.Start - para.Start + 1, textRange.Length - 1));
						}
						if (!Regex.IsMatch(textRange.Text, AH.A(14628)))
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
						if (!string.IsNullOrEmpty(textRange.Font.Name))
						{
							list.Add(textRange.Font.Name);
						}
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_0112;
						}
						continue;
						end_IL_0112:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (5)
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
					List<string> source = list;
					Func<string, S<string>> keySelector;
					if (_Closure_0024__.A == null)
					{
						keySelector = (_Closure_0024__.A = [SpecialName] (string A) => new S<string>(A.ToString()));
					}
					else
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
						keySelector = _Closure_0024__.A;
					}
					enumerator2 = source.GroupBy(keySelector).GetEnumerator();
					while (enumerator2.MoveNext())
					{
						IGrouping<S<string>, string> current = enumerator2.Current;
						list2.Add(current.Key.Font);
						list3.Add(current.Key.Font + AH.A(14248) + current.Count() + AH.A(14255));
						current = null;
					}
				}
				finally
				{
					if (enumerator2 != null)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							enumerator2.Dispose();
							break;
						}
					}
				}
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.MultipleFontFamilies(sld, shp, para, list3, string.Join(AH.A(14258), list3.ToArray()), list2));
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
