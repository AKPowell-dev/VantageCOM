using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.UI;

public sealed class Highlight
{
	private static readonly int m_A = ColorTranslator.ToOle(Color.FromArgb(210, 240, 224));

	private static readonly int m_B = ColorTranslator.ToOle(Color.Transparent);

	private static readonly string m_A = VH.A(36227);

	internal static void A(Range A)
	{
		FormatConditions formatConditions = A.FormatConditions;
		formatConditions.Add(XlFormatConditionType.xlExpression, RuntimeHelpers.GetObjectValue(Missing.Value), Highlight.m_A, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		object instance = NewLateBinding.LateGet(formatConditions.Item(formatConditions.Count), null, VH.A(36170), new object[0], null, null, null);
		NewLateBinding.LateSetComplex(instance, null, VH.A(36187), new object[1] { Highlight.m_A }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(instance, null, VH.A(36212), new object[1] { XlPattern.xlPatternGray50 }, null, null, OptimisticSet: false, RValueBase: true);
		instance = null;
		_ = null;
	}

	internal static void B(Range A)
	{
		FormatConditions formatConditions = A.FormatConditions;
		for (int i = formatConditions.Count; i >= 1; i = checked(i + -1))
		{
			try
			{
				FormatCondition formatCondition = (FormatCondition)formatConditions.Item(i);
				if (formatCondition.Type == 2 && Operators.CompareString(formatCondition.Formula1, Highlight.m_A, TextCompare: false) == 0)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							formatCondition.Delete();
							goto end_IL_007a;
						}
					}
				}
				formatCondition = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			continue;
			end_IL_007a:
			break;
		}
		formatConditions = null;
	}
}
