using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class Tables
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<CellMargins, U<float, float, float, float>> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal U<float, float, float, float> A(CellMargins A)
		{
			return new U<float, float, float, float>(A.Top, A.Right, A.Bottom, A.Left);
		}
	}

	public static void CellMargins(Slide sld, Shape shp, List<CellMargins> listMargins, List<Shape> listShapes)
	{
		Func<CellMargins, U<float, float, float, float>> keySelector;
		if (_Closure_0024__.A == null)
		{
			keySelector = (_Closure_0024__.A = [SpecialName] (CellMargins A) => new U<float, float, float, float>(A.Top, A.Right, A.Bottom, A.Left));
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
		IEnumerable<IGrouping<U<float, float, float, float>, CellMargins>> enumerable = listMargins.GroupBy(keySelector);
		if (enumerable.Count() > 1)
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
			List<CellMargins> list = new List<CellMargins>();
			List<string> list2 = new List<string>();
			IEnumerator<IGrouping<U<float, float, float, float>, CellMargins>> enumerator = default(IEnumerator<IGrouping<U<float, float, float, float>, CellMargins>>);
			try
			{
				enumerator = enumerable.GetEnumerator();
				while (enumerator.MoveNext())
				{
					IGrouping<U<float, float, float, float>, CellMargins> current = enumerator.Current;
					CellMargins item = default(CellMargins);
					U<float, float, float, float> key = current.Key;
					item.Top = key.Top;
					item.Right = key.Right;
					item.Bottom = key.Bottom;
					item.Left = key.Left;
					list2.Add(Conversions.ToString(key.Top) + AH.A(14600) + Conversions.ToString(key.Right) + AH.A(14600) + Conversions.ToString(key.Bottom) + AH.A(14600) + Conversions.ToString(key.Left) + AH.A(14611) + current.Count() + AH.A(14255));
					key = null;
					list.Add(item);
					item = default(CellMargins);
				}
			}
			finally
			{
				if (enumerator != null)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						enumerator.Dispose();
						break;
					}
				}
			}
			Main.Analysis.Errors.Add(new TableCellMargins(sld, shp, list2, listShapes, list));
			list = null;
			list2 = null;
		}
		enumerable = null;
	}
}
