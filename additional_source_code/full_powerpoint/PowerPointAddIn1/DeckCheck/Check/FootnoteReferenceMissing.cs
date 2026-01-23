using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class FootnoteReferenceMissing
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<int, int> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal int A(int A)
		{
			return A;
		}
	}

	public void Check(Slide sld, List<int> listFootnoteNumbers, List<int> listFoundFootnotes, Shape shp)
	{
		if (listFootnoteNumbers.Count <= 0)
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
			listFootnoteNumbers = listFootnoteNumbers.Distinct().ToList();
			List<int> source = listFootnoteNumbers;
			Func<int, int> keySelector;
			if (_Closure_0024__.A == null)
			{
				keySelector = (_Closure_0024__.A = [SpecialName] (int A) => A);
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
				keySelector = _Closure_0024__.A;
			}
			listFootnoteNumbers = source.OrderBy(keySelector).ToList();
			using List<int>.Enumerator enumerator = listFootnoteNumbers.GetEnumerator();
			while (enumerator.MoveNext())
			{
				int current = enumerator.Current;
				if (listFoundFootnotes.Contains(current))
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
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.FootnoteReferenceMissing(sld, shp, current));
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
	}
}
