using System;
using System.Collections;
using A;
using ExcelAddIn1.SuperFind2.Results;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class Font
{
	internal static void A(WorksheetItem A, Range B)
	{
		Helpers.A(A, B, Font.A);
	}

	private static bool A(Range A)
	{
		return Conversions.ToBoolean(A.Font.Bold);
	}

	internal static void B(WorksheetItem A, Range B)
	{
		Helpers.A(A, B, Font.B);
	}

	private static bool B(Range A)
	{
		return !Conversions.ToBoolean(A.Font.Bold);
	}

	internal static void C(WorksheetItem A, Range B)
	{
		Helpers.A(A, B, C);
	}

	private static bool C(Range A)
	{
		return Conversions.ToBoolean(A.Font.Italic);
	}

	internal static void D(WorksheetItem A, Range B)
	{
		Helpers.A(A, B, D);
	}

	private static bool D(Range A)
	{
		return !Conversions.ToBoolean(A.Font.Italic);
	}

	internal static void E(WorksheetItem A, Range B)
	{
		Helpers.A(A, B, E);
	}

	private static bool E(Range A)
	{
		return Operators.ConditionalCompareObjectNotEqual(A.Font.Underline, XlUnderlineStyle.xlUnderlineStyleNone, TextCompare: false);
	}

	internal static void F(WorksheetItem A, Range B)
	{
		Helpers.A(A, B, F);
	}

	private static bool F(Range A)
	{
		return Operators.ConditionalCompareObjectEqual(A.Font.Underline, XlUnderlineStyle.xlUnderlineStyleNone, TextCompare: false);
	}

	internal static void G(WorksheetItem A, Range B)
	{
		Font.A(A, B, (Func<float, float, bool>)Font.A);
	}

	private static bool A(float A, float B)
	{
		return A == B;
	}

	internal static void H(WorksheetItem A, Range B)
	{
		Font.A(A, B, (Func<float, float, bool>)Font.B);
	}

	private static bool B(float A, float B)
	{
		return A != B;
	}

	internal static void I(WorksheetItem A, Range B)
	{
		Font.A(A, B, C);
	}

	private static bool C(float A, float B)
	{
		return A > B;
	}

	internal static void J(WorksheetItem A, Range B)
	{
		Font.A(A, B, D);
	}

	private static bool D(float A, float B)
	{
		return A < B;
	}

	private static void A(WorksheetItem A, Range B, Func<float, float, bool> C)
	{
		Range A2 = null;
		float arg;
		try
		{
			arg = Conversions.ToSingle(Props.SearchForm.Input1);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.WarningMessage(VH.A(103134));
			ProjectData.ClearProjectError();
			return;
		}
		if (!Props.SearchForm.LookInEmptyCells)
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
			B = RangeHelpers.H(B);
		}
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			try
			{
				enumerator = B.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					try
					{
						if (C(Conversions.ToSingle(range.Font.Size), arg))
						{
							RangeHelpers.A(ref A2, range);
						}
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_00cf;
					}
					continue;
					end_IL_00cf:
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			Helpers.B(A, A2);
			A2 = null;
			return;
		}
	}

	internal static void K(WorksheetItem A, Range B)
	{
		Font.A(A, B, (Func<float, float, float, bool>)Font.A);
	}

	private static bool A(float A, float B, float C)
	{
		if (A > B)
		{
			return A < C;
		}
		return false;
	}

	internal static void L(WorksheetItem A, Range B)
	{
		Font.A(A, B, (Func<float, float, float, bool>)Font.B);
	}

	private static bool B(float A, float B, float C)
	{
		if (!(A <= B))
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return A >= C;
				}
			}
		}
		return true;
	}

	private static void A(WorksheetItem A, Range B, Func<float, float, float, bool> C)
	{
		Range A2 = null;
		float arg;
		float arg2;
		try
		{
			arg = Conversions.ToSingle(Props.SearchForm.Input1);
			arg2 = Conversions.ToSingle(Props.SearchForm.Input2);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.WarningMessage(VH.A(103134));
			ProjectData.ClearProjectError();
			return;
		}
		if (!Props.SearchForm.LookInEmptyCells)
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
			B = RangeHelpers.H(B);
		}
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			try
			{
				enumerator = B.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					try
					{
						if (C(Conversions.ToSingle(range.Font.Size), arg, arg2))
						{
							RangeHelpers.A(ref A2, range);
						}
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_00e9;
					}
					continue;
					end_IL_00e9:
					break;
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
			Helpers.B(A, A2);
			A2 = null;
			return;
		}
	}

	internal static void M(WorksheetItem A, Range B)
	{
		Font.A(A, B, E);
	}

	private static bool E(string A, string B)
	{
		return Operators.CompareString(A, B, TextCompare: false) == 0;
	}

	internal static void N(WorksheetItem A, Range B)
	{
		Font.A(A, B, F);
	}

	private static bool F(string A, string B)
	{
		return Operators.CompareString(A, B, TextCompare: false) != 0;
	}

	private static void A(WorksheetItem A, Range B, Func<string, string, bool> C)
	{
		Range A2 = null;
		string arg = Props.SearchForm.Input1.ToLower();
		if (!Props.SearchForm.LookInEmptyCells)
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
			B = RangeHelpers.H(B);
		}
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = B.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range range = (Range)enumerator.Current;
				try
				{
					if (!C(range.Font.Name.ToString().ToLower(), arg))
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
						RangeHelpers.A(ref A2, range);
						break;
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
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		Helpers.B(A, A2);
		A2 = null;
	}
}
