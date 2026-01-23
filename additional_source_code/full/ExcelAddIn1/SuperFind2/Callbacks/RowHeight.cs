using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using ExcelAddIn1.SuperFind2.Results;
using ExcelAddIn1.SuperFind2.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class RowHeight
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<string, string, bool> A;

		public static Func<string, string, bool> B;

		public static Func<string, string, bool> C;

		public static Func<string, string, bool> D;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(string A, string B)
		{
			return RowHeight.A(A, Conversions.ToSingle(B));
		}

		[SpecialName]
		internal bool B(string A, string B)
		{
			return RowHeight.B(A, Conversions.ToSingle(B));
		}

		[SpecialName]
		internal bool C(string A, string B)
		{
			return RowHeight.C((object)A, Conversions.ToSingle(B));
		}

		[SpecialName]
		internal bool D(string A, string B)
		{
			return RowHeight.D((object)A, Conversions.ToSingle(B));
		}
	}

	internal static void A(WorksheetItem A, Range B)
	{
		Func<string, string, bool> c;
		if (_Closure_0024__.A == null)
		{
			c = (_Closure_0024__.A = [SpecialName] (string a, string value) => RowHeight.A(a, Conversions.ToSingle(value)));
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			c = _Closure_0024__.A;
		}
		RowHeight.A(A, B, c);
	}

	private static bool A(object A, float B)
	{
		return Operators.ConditionalCompareObjectEqual(A, B, TextCompare: false);
	}

	internal static void B(WorksheetItem A, Range B)
	{
		Func<string, string, bool> c;
		if (_Closure_0024__.B == null)
		{
			c = (_Closure_0024__.B = [SpecialName] (string a, string value) => RowHeight.B(a, Conversions.ToSingle(value)));
		}
		else
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
			c = _Closure_0024__.B;
		}
		RowHeight.A(A, B, c);
	}

	private static bool B(object A, float B)
	{
		return Operators.ConditionalCompareObjectEqual(A, B, TextCompare: false);
	}

	internal static void C(WorksheetItem A, Range B)
	{
		Func<string, string, bool> c;
		if (_Closure_0024__.C == null)
		{
			c = (_Closure_0024__.C = [SpecialName] (string a, string value) => C(a, Conversions.ToSingle(value)));
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			c = _Closure_0024__.C;
		}
		RowHeight.A(A, B, c);
	}

	private static bool C(object A, float B)
	{
		return Operators.ConditionalCompareObjectGreater(A, B, TextCompare: false);
	}

	internal static void D(WorksheetItem A, Range B)
	{
		Func<string, string, bool> c;
		if (_Closure_0024__.D == null)
		{
			c = (_Closure_0024__.D = [SpecialName] (string a, string value) => D(a, Conversions.ToSingle(value)));
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
			c = _Closure_0024__.D;
		}
		RowHeight.A(A, B, c);
	}

	private static bool D(object A, float B)
	{
		return Operators.ConditionalCompareObjectLess(A, B, TextCompare: false);
	}

	private static void A(WorksheetItem A, Range B, Func<string, string, bool> C)
	{
		string input = Props.SearchForm.Input1;
		B = (Range)B.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
		if (!Props.SearchForm.LookInEmptyCells)
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
			B = RangeHelpers.H(B);
		}
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			float num = Conversions.ToSingle(input);
			try
			{
				enumerator = B.Cells.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					try
					{
						if (C(Conversions.ToString(range.RowHeight), Conversions.ToString(num)))
						{
							A.D(range.EntireRow);
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (5)
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
		}
	}
}
