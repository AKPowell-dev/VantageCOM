using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

internal sealed class LG
{
	[CompilerGenerated]
	private int m_A;

	[CompilerGenerated]
	private int m_B;

	[CompilerGenerated]
	private object m_A;

	[CompilerGenerated]
	private double m_A;

	private int Color
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

	private int ColorIndex
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	private object Pattern
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

	private double PatternColor
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

	internal LG()
	{
	}

	internal void A(ref Interior A)
	{
		Interior interior = A;
		Color = Conversions.ToInteger(interior.Color);
		ColorIndex = Conversions.ToInteger(interior.ColorIndex);
		Pattern = RuntimeHelpers.GetObjectValue(interior.Pattern);
		PatternColor = Conversions.ToDouble(interior.PatternColor);
		interior = null;
	}

	internal void B(ref Interior A)
	{
		Interior interior = A;
		if (ColorIndex == -4142)
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
			interior.ColorIndex = Constants.xlNone;
			interior.Pattern = XlPattern.xlPatternNone;
		}
		else
		{
			interior.Color = Color;
			interior.Pattern = RuntimeHelpers.GetObjectValue(Pattern);
			interior.PatternColor = PatternColor;
		}
		interior = null;
	}
}
