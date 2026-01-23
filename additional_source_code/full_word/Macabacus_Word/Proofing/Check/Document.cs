using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Errors;

namespace Macabacus_Word.Proofing.Check;

public sealed class Document
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<string, C<string>> A;

		public static Func<float, D<float>> A;

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
		internal D<float> A(float A)
		{
			return new D<float>(A);
		}
	}

	public static void GlobalFontFamilies(List<string> listFontFamilies, Settings options)
	{
		List<string> list = new List<string>();
		using (IEnumerator<IGrouping<C<string>, string>> enumerator = (from A in listFontFamilies
			group A by new C<string>(A.ToString())).GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				IGrouping<C<string>, string> current = enumerator.Current;
				list.Add(current.Key.Font + XC.A(22691) + current.Count() + XC.A(20696));
				current = null;
			}
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
				break;
			}
		}
		string strSubtitle = string.Join(XC.A(22698), list.ToArray());
		if (list.Count > options.MaxFontFamilies)
		{
			Main.Analysis.Errors.Add(new FontFamilyWarning(strSubtitle));
		}
		list = null;
	}

	public static void GlobalFontSizes(List<float> listFontSizes, Settings options)
	{
		List<string> list = new List<string>();
		IEnumerator<IGrouping<D<float>, float>> enumerator = default(IEnumerator<IGrouping<D<float>, float>>);
		try
		{
			Func<float, D<float>> keySelector;
			if (_Closure_0024__.A == null)
			{
				keySelector = (_Closure_0024__.A = [SpecialName] (float A) => new D<float>(A));
			}
			else
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
				keySelector = _Closure_0024__.A;
			}
			enumerator = listFontSizes.GroupBy(keySelector).GetEnumerator();
			while (enumerator.MoveNext())
			{
				IGrouping<D<float>, float> current = enumerator.Current;
				list.Add(current.Key.Size + XC.A(22703) + current.Count() + XC.A(20696));
				current = null;
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_00ae;
				}
				continue;
				end_IL_00ae:
				break;
			}
		}
		finally
		{
			enumerator?.Dispose();
		}
		string strSubtitle = string.Join(XC.A(22698), list.ToArray());
		if (list.Count > ((Settings)options).MaxFontSizes)
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
			Main.Analysis.Errors.Add(new FontSizeWarning(strSubtitle));
		}
		list = null;
	}
}
