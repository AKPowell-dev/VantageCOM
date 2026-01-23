using System;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Format;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class CycleColor
{
	private static string m_A;

	private static Chart m_A;

	public static void Cycle()
	{
		if (!Helpers.A())
		{
			return;
		}
		checked
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Application application = MH.A.Application;
				if (application.ActiveChart != null)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					object objectValue = RuntimeHelpers.GetObjectValue(application.Selection);
					string text;
					try
					{
						text = Conversions.ToString(NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null));
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						text = "";
						ProjectData.ClearProjectError();
					}
					ColorCycle chartColorCycle = KH.A.ChartColorCycle;
					if (CycleColor.m_A == application.ActiveChart && Operators.CompareString(text, CycleColor.m_A, TextCompare: false) == 0)
					{
						if (chartColorCycle.Index == chartColorCycle.Colors.Count)
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
							chartColorCycle.Index = 0;
						}
					}
					else
					{
						chartColorCycle.Index = 0;
					}
					A(RuntimeHelpers.GetObjectValue(objectValue), chartColorCycle.Colors[chartColorCycle.Index].OLE);
					CycleColor.m_A = application.ActiveChart;
					CycleColor.m_A = text;
					chartColorCycle.Index++;
					if (chartColorCycle.Index == 1)
					{
						clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, chartColorCycle.Activity);
					}
					chartColorCycle = null;
					objectValue = null;
				}
				application = null;
				return;
			}
		}
	}

	public static void DoChartColor(IRibbonControl control)
	{
		Application application = MH.A.Application;
		if (application.ActiveChart != null)
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
			A(RuntimeHelpers.GetObjectValue(application.Selection), clsColors.RGB2Ole(control.Tag));
		}
		application = null;
	}

	private static void A(object A, int B)
	{
		try
		{
			if (A is Series)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
					{
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						CycleColor.A(((Series)A).Format, B);
						LineFormat line = ((Series)A).Format.Line;
						if (line.Visible == MsoTriState.msoTrue)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								break;
							}
							if (line.Weight > 0f)
							{
								line.ForeColor.RGB = B;
								line.BackColor.RGB = B;
							}
						}
						line = null;
						return;
					}
					}
				}
			}
			if (A is Point)
			{
				CycleColor.A(((Point)A).Format, B);
				return;
			}
			if (A is PlotArea)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						CycleColor.A(((PlotArea)A).Format, B);
						return;
					}
				}
			}
			if (A is ChartArea)
			{
				CycleColor.A(((ChartArea)A).Format, B);
				return;
			}
			if (A is Legend)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						CycleColor.A(((Legend)A).Format, B);
						return;
					}
				}
			}
			if (!(A is Gridlines))
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
				LineFormat line2 = ((Gridlines)A).Format.Line;
				line2.ForeColor.RGB = B;
				line2.BackColor.RGB = B;
				_ = null;
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	private static void A(ChartFormat A, int B)
	{
		Microsoft.Office.Interop.Excel.FillFormat fill = A.Fill;
		fill.ForeColor.RGB = B;
		fill.BackColor.RGB = B;
		_ = null;
	}
}
