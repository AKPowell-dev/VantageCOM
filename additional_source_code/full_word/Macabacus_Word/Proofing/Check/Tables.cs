using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.Check;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class Tables
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<CellPadding, E<float, float, float, float>> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal E<float, float, float, float> A(CellPadding A)
		{
			return new E<float, float, float, float>(A.Top, A.Right, A.Bottom, A.Left);
		}
	}

	public static void CellMargins(Table tbl, List<CellPadding> listMargins, List<Range> listRanges)
	{
		Func<CellPadding, E<float, float, float, float>> keySelector;
		if (_Closure_0024__.A == null)
		{
			keySelector = (_Closure_0024__.A = [SpecialName] (CellPadding A) => new E<float, float, float, float>(A.Top, A.Right, A.Bottom, A.Left));
		}
		else
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
			keySelector = _Closure_0024__.A;
		}
		IEnumerable<IGrouping<E<float, float, float, float>, CellPadding>> enumerable = listMargins.GroupBy(keySelector);
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
			List<CellPadding> list = new List<CellPadding>();
			List<string> list2 = new List<string>();
			IEnumerator<IGrouping<E<float, float, float, float>, CellPadding>> enumerator = default(IEnumerator<IGrouping<E<float, float, float, float>, CellPadding>>);
			try
			{
				enumerator = enumerable.GetEnumerator();
				while (enumerator.MoveNext())
				{
					IGrouping<E<float, float, float, float>, CellPadding> current = enumerator.Current;
					CellPadding item = default(CellPadding);
					E<float, float, float, float> key = current.Key;
					item.Top = key.Top;
					item.Right = key.Right;
					item.Bottom = key.Bottom;
					item.Left = key.Left;
					list2.Add(Conversions.ToString(key.Top) + XC.A(24507) + Conversions.ToString(key.Right) + XC.A(24507) + Conversions.ToString(key.Bottom) + XC.A(24507) + Conversions.ToString(key.Left) + XC.A(22703) + current.Count() + XC.A(20696));
					key = null;
					list.Add(item);
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_01a2;
					}
					continue;
					end_IL_01a2:
					break;
				}
			}
			finally
			{
				if (enumerator != null)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						enumerator.Dispose();
						break;
					}
				}
			}
			Main.Analysis.Errors.Add(new TableCellMargins(tbl, list2, listRanges, list));
			list = null;
			list2 = null;
		}
		enumerable = null;
	}

	public static void Colors(Table tbl, List<int> listColors, Severity sev)
	{
		//IL_0103: Unknown result type (might be due to invalid IL or missing references)
		//IL_0172: Unknown result type (might be due to invalid IL or missing references)
		Dictionary<int, List<Range>> dictionary = new Dictionary<int, List<Range>>();
		Dictionary<int, List<Range>> D = new Dictionary<int, List<Range>>();
		try
		{
			Table table = tbl;
			int rGB = table.Range.Font.TextColor.RGB;
			if (Color.ColorNotInPalette(rGB, listColors))
			{
				List<Range> value = null;
				if (dictionary.TryGetValue(rGB, out value))
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
					value.Add(table.Range);
					dictionary[rGB] = value;
				}
				else
				{
					value = new List<Range>();
					value.Add(table.Range);
					dictionary.Add(rGB, value);
				}
				value = null;
			}
			A(table.Shading.BackgroundPatternColor, table.Range, listColors, ref D);
			A(table.Shading.ForegroundPatternColor, table.Range, listColors, ref D);
			if (dictionary.Count > 0)
			{
				foreach (KeyValuePair<int, List<Range>> item in dictionary)
				{
					Main.Analysis.Errors.Add(new NonconformingTableCellFontColor(tbl, item.Key, item.Value, sev));
				}
			}
			if (D.Count > 0)
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
				using Dictionary<int, List<Range>>.Enumerator enumerator2 = D.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					KeyValuePair<int, List<Range>> current2 = enumerator2.Current;
					Main.Analysis.Errors.Add(new NonconformingTableCellFillColor(tbl, current2.Key, current2.Value, sev));
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_0188;
					}
					continue;
					end_IL_0188:
					break;
				}
			}
			table = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Interaction.MsgBox(ex2.Message);
			ProjectData.ClearProjectError();
		}
		dictionary = null;
		D = null;
	}

	private static void A(WdColor A, Range B, List<int> C, ref Dictionary<int, List<Range>> D)
	{
		int num = Tables.A(A);
		if (!Color.ColorNotInPalette(num, C))
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
			List<Range> value = null;
			if (D.TryGetValue(num, out value))
			{
				value.Add(B);
				D[num] = value;
			}
			else
			{
				value = new List<Range>();
				value.Add(B);
				D.Add(num, value);
			}
			value = null;
			return;
		}
	}

	private static int A(WdColor A)
	{
		return TableColors.Colours2(A);
	}

	private static WdColor A(Color A)
	{
		return (WdColor)Information.RGB(A.R, A.G, A.B);
	}
}
