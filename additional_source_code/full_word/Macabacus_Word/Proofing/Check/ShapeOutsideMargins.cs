using System;
using System.Runtime.CompilerServices;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class ShapeOutsideMargins
{
	private double m_A;

	private double m_B;

	private double m_C;

	private double m_D;

	private double A
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	private double B
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
		}
	}

	private double C
	{
		get
		{
			return this.m_C;
		}
		set
		{
			this.m_C = value;
		}
	}

	private double D
	{
		get
		{
			return this.m_D;
		}
		set
		{
			this.m_D = value;
		}
	}

	public ShapeOutsideMargins(double dblLeftMargin, double dblRightMargin, double dblTopMargin, double dblBottomMargin)
	{
		A = dblLeftMargin;
		B = dblRightMargin;
		C = dblTopMargin;
		D = dblBottomMargin;
	}

	public void Check(Microsoft.Office.Interop.Word.Document doc, Shape shp)
	{
		try
		{
			Check(doc, shp, shp.Top, shp.Left, shp.Height, shp.Width);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public void Check(Microsoft.Office.Interop.Word.Document doc, InlineShape shp)
	{
		try
		{
			doc.ActiveWindow.GetPoint(out var ScreenPixelsLeft, out var ScreenPixelsTop, out var _, out var _, shp);
			Check(doc, shp, ScreenPixelsTop, ScreenPixelsLeft, shp.Height, shp.Width);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public void Check(Microsoft.Office.Interop.Word.Document doc, object shp, float sngTop, float sngLeft, float sngHeight, float sngWidth)
	{
		if (!(A > 0.0))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (!(B > 0.0))
			{
				return;
			}
			try
			{
				double num = Math.Round(sngLeft, 4);
				double num2 = Math.Round(sngLeft + sngWidth, 4);
				double num3 = Math.Round(sngTop, 4);
				double num4 = Math.Round(sngTop + sngHeight, 4);
				if (num < A)
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
					if (num >= 0.0)
					{
						goto IL_00d0;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						break;
					}
				}
				if (num2 > B)
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
					if (num2 <= Math.Round(doc.PageSetup.PageWidth, 4))
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
						goto IL_00d0;
					}
				}
				if (num3 < C)
				{
					if (num3 >= 0.0)
					{
						goto IL_014c;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						break;
					}
				}
				if (!(num4 > D))
				{
					return;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				if (!(num4 <= Math.Round(doc.PageSetup.PageHeight, 4)))
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
					break;
				}
				goto IL_014c;
				IL_00d0:
				Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.ShapeOutsideMargins(RuntimeHelpers.GetObjectValue(shp)));
				return;
				IL_014c:
				Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.ShapeOutsideMargins(RuntimeHelpers.GetObjectValue(shp)));
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
