using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

internal sealed class JG
{
	[CompilerGenerated]
	private XlLineStyle m_A;

	[CompilerGenerated]
	private XlBorderWeight m_A;

	[CompilerGenerated]
	private object m_A;

	private XlLineStyle Style
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private XlBorderWeight Weight
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private object Color
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = RuntimeHelpers.GetObjectValue(value);
		}
	}

	internal JG()
	{
	}

	internal void A(ref Borders A, XlBordersIndex B)
	{
		Border border = A[B];
		Style = (XlLineStyle)Conversions.ToInteger(border.LineStyle);
		Weight = (XlBorderWeight)Conversions.ToInteger(border.Weight);
		Color = RuntimeHelpers.GetObjectValue(border.Color);
		border = null;
	}

	internal void B(ref Borders A, XlBordersIndex B)
	{
		Border border = A[B];
		border.LineStyle = Style;
		if (Operators.ConditionalCompareObjectNotEqual(border.LineStyle, Constants.xlNone, TextCompare: false))
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
			border.Weight = Weight;
			border.Color = RuntimeHelpers.GetObjectValue(Color);
		}
		border = null;
	}
}
