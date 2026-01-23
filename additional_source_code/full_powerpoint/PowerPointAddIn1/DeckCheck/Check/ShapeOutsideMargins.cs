using System;
using System.Runtime.CompilerServices;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class ShapeOutsideMargins
{
	[CompilerGenerated]
	private double m_A;

	[CompilerGenerated]
	private double m_B;

	private const int m_A = 4;

	private double LeftSlideMargin
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

	private double RightSlideMargin
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

	public ShapeOutsideMargins(double dblLeftMargin, double dblRightMargin)
	{
		LeftSlideMargin = dblLeftMargin;
		RightSlideMargin = dblRightMargin;
	}

	public void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		if (shp.Type == MsoShapeType.msoPlaceholder)
		{
			return;
		}
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
			if (!(LeftSlideMargin > 0.0) || !(RightSlideMargin > 0.0))
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
				try
				{
					double num = shp.Left;
					double num2 = shp.Left + shp.Width;
					if (A(num, LeftSlideMargin))
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
						if (B(0.0, num))
						{
							goto IL_00de;
						}
					}
					if (!A(RightSlideMargin, num2))
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
						break;
					}
					if (!B(num2, sld.CustomLayout.Width))
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
						break;
					}
					goto IL_00de;
					IL_00de:
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.ShapeOutsideMargins(sld, shp));
					return;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
					return;
				}
			}
		}
	}

	private static bool A(double A, double B)
	{
		return modFunctionsNum.RoundUp(A, 4) < modFunctionsNum.Truncate(B, 4);
	}

	private static bool B(double A, double B)
	{
		return modFunctionsNum.RoundUp(A, 4) <= modFunctionsNum.Truncate(B, 4);
	}
}
