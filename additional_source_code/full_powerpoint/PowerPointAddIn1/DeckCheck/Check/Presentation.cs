using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class Presentation
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<float, T<float>> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal T<float> A(float A)
		{
			return new T<float>(A);
		}
	}

	public static void GlobalFontSizes(List<float> listFontSizes, Settings options)
	{
		List<string> list = new List<string>();
		IEnumerator<IGrouping<T<float>, float>> enumerator = default(IEnumerator<IGrouping<T<float>, float>>);
		try
		{
			Func<float, T<float>> keySelector;
			if (_Closure_0024__.A == null)
			{
				keySelector = (_Closure_0024__.A = [SpecialName] (float A) => new T<float>(A));
			}
			else
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
				keySelector = _Closure_0024__.A;
			}
			enumerator = listFontSizes.GroupBy(keySelector).GetEnumerator();
			while (enumerator.MoveNext())
			{
				IGrouping<T<float>, float> current = enumerator.Current;
				list.Add(current.Key.Size + AH.A(14611) + current.Count() + AH.A(14255));
				current = null;
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_00b2;
				}
				continue;
				end_IL_00b2:
				break;
			}
		}
		finally
		{
			enumerator?.Dispose();
		}
		string strSubtitle = string.Join(AH.A(14258), list.ToArray());
		if (list.Count > ((Settings)options).MaxFontSizes)
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
			Main.Analysis.Errors.Add(new FontSizeWarning(strSubtitle));
		}
		list = null;
	}
}
