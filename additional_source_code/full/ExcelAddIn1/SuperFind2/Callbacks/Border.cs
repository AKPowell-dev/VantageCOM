using System;
using System.Collections;
using ExcelAddIn1.Format;
using ExcelAddIn1.SuperFind2.Results;
using ExcelAddIn1.SuperFind2.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class Border
{
	internal static void A(WorksheetItem A, Range B)
	{
		Border.A(A, B, Border.A);
	}

	private static bool A(Microsoft.Office.Interop.Excel.Borders A)
	{
		if (Operators.ConditionalCompareObjectNotEqual(A[XlBordersIndex.xlEdgeTop].LineStyle, Constants.xlNone, TextCompare: false))
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return true;
				}
			}
		}
		if (Operators.ConditionalCompareObjectNotEqual(A[XlBordersIndex.xlEdgeBottom].LineStyle, Constants.xlNone, TextCompare: false))
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		if (Operators.ConditionalCompareObjectNotEqual(A[XlBordersIndex.xlEdgeLeft].LineStyle, Constants.xlNone, TextCompare: false))
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		if (Operators.ConditionalCompareObjectNotEqual(A[XlBordersIndex.xlEdgeRight].LineStyle, Constants.xlNone, TextCompare: false))
		{
			return true;
		}
		if (Operators.ConditionalCompareObjectNotEqual(A[XlBordersIndex.xlDiagonalDown].LineStyle, Constants.xlNone, TextCompare: false))
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		if (Operators.ConditionalCompareObjectNotEqual(A[XlBordersIndex.xlDiagonalUp].LineStyle, Constants.xlNone, TextCompare: false))
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		return false;
	}

	internal static void B(WorksheetItem A, Range B)
	{
		Border.A(A, B, Border.B);
	}

	private static bool B(Microsoft.Office.Interop.Excel.Borders A)
	{
		return ExcelAddIn1.Format.Borders.HasTopBorder(A);
	}

	internal static void C(WorksheetItem A, Range B)
	{
		Border.A(A, B, C);
	}

	private static bool C(Microsoft.Office.Interop.Excel.Borders A)
	{
		return ExcelAddIn1.Format.Borders.HasBottomBorder(A);
	}

	internal static void D(WorksheetItem A, Range B)
	{
		Border.A(A, B, D);
	}

	private static bool D(Microsoft.Office.Interop.Excel.Borders A)
	{
		return ExcelAddIn1.Format.Borders.HasLeftBorder(A);
	}

	internal static void E(WorksheetItem A, Range B)
	{
		Border.A(A, B, E);
	}

	private static bool E(Microsoft.Office.Interop.Excel.Borders A)
	{
		return ExcelAddIn1.Format.Borders.HasRightBorder(A);
	}

	private static void A(WorksheetItem A, Range B, Func<Microsoft.Office.Interop.Excel.Borders, bool> C)
	{
		Range A2 = null;
		if (!Props.SearchForm.LookInEmptyCells)
		{
			B = RangeHelpers.H(B);
		}
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			try
			{
				enumerator = B.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					try
					{
						if (!C(range.Borders))
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
							RangeHelpers.A(ref A2, range);
							break;
						}
						continue;
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
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_008c;
					}
					continue;
					end_IL_008c:
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
}
