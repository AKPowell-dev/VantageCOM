using ExcelAddIn1.SuperFind2.Results;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class Alignment
{
	internal static void A(WorksheetItem A, Range B)
	{
		Helpers.A(A, B, Alignment.A);
	}

	private static bool A(Range A)
	{
		return Operators.ConditionalCompareObjectEqual(A.VerticalAlignment, XlVAlign.xlVAlignTop, TextCompare: false);
	}

	internal static void B(WorksheetItem A, Range B)
	{
		Helpers.A(A, B, Alignment.B);
	}

	private static bool B(Range A)
	{
		return Operators.ConditionalCompareObjectEqual(A.VerticalAlignment, XlVAlign.xlVAlignBottom, TextCompare: false);
	}

	internal static void C(WorksheetItem A, Range B)
	{
		Helpers.A(A, B, C);
	}

	private static bool C(Range A)
	{
		return Operators.ConditionalCompareObjectEqual(A.VerticalAlignment, XlVAlign.xlVAlignCenter, TextCompare: false);
	}

	internal static void D(WorksheetItem A, Range B)
	{
		Helpers.A(A, B, D);
	}

	private static bool D(Range A)
	{
		return Operators.ConditionalCompareObjectEqual(A.HorizontalAlignment, XlHAlign.xlHAlignLeft, TextCompare: false);
	}

	internal static void E(WorksheetItem A, Range B)
	{
		Helpers.A(A, B, E);
	}

	private static bool E(Range A)
	{
		return Operators.ConditionalCompareObjectEqual(A.HorizontalAlignment, XlHAlign.xlHAlignCenter, TextCompare: false);
	}

	internal static void F(WorksheetItem A, Range B)
	{
		Helpers.A(A, B, F);
	}

	private static bool F(Range A)
	{
		return Operators.ConditionalCompareObjectEqual(A.HorizontalAlignment, XlHAlign.xlHAlignCenterAcrossSelection, TextCompare: false);
	}

	internal static void G(WorksheetItem A, Range B)
	{
		Helpers.A(A, B, G);
	}

	private static bool G(Range A)
	{
		return Operators.ConditionalCompareObjectEqual(A.HorizontalAlignment, XlHAlign.xlHAlignLeft, TextCompare: false);
	}

	internal static void H(WorksheetItem A, Range B)
	{
		Helpers.A(A, B, H);
	}

	private static bool H(Range A)
	{
		int num;
		if (Conversions.ToBoolean(Operators.CompareObjectEqual(A.HorizontalAlignment, XlHAlign.xlHAlignLeft, TextCompare: false)))
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
			num = (Conversions.ToBoolean(Operators.CompareObjectGreater(A.IndentLevel, 0, TextCompare: false)) ? 1 : 0);
		}
		else
		{
			num = 0;
		}
		return Conversions.ToBoolean((byte)num != 0);
	}

	internal static void I(WorksheetItem A, Range B)
	{
		Helpers.A(A, B, I);
	}

	private static bool I(Range A)
	{
		int num;
		if (Conversions.ToBoolean(Operators.CompareObjectEqual(A.HorizontalAlignment, XlHAlign.xlHAlignRight, TextCompare: false)))
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
			num = (Conversions.ToBoolean(Operators.CompareObjectGreater(A.IndentLevel, 0, TextCompare: false)) ? 1 : 0);
		}
		else
		{
			num = 0;
		}
		return Conversions.ToBoolean((byte)num != 0);
	}
}
