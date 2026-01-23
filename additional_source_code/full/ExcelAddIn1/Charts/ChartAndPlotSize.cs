using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class ChartAndPlotSize
{
	public static void ShowDialog()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		wpfChartSize wpfChartSize2 = default(wpfChartSize);
		Chart chart = default(Chart);
		Application application = default(Application);
		Worksheet worksheet = default(Worksheet);
		ChartObject chartObject = default(ChartObject);
		Pane pane = default(Pane);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					if (!Helpers.A())
					{
						goto end_IL_0000;
					}
					while (true)
					{
						switch (7)
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
					goto IL_0021;
				case 1212:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000_2;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 3:
							goto IL_0021;
						case 4:
							goto IL_0028;
						case 5:
							goto IL_0032;
						case 6:
							goto IL_003a;
						case 7:
							goto IL_0048;
						case 8:
							goto IL_005a;
						case 9:
							goto IL_00a9;
						case 10:
							goto IL_00c2;
						case 11:
							goto IL_00d7;
						case 12:
							goto IL_039f;
						case 13:
							goto IL_03ac;
						case 14:
							goto IL_03c5;
						case 15:
							goto IL_03d6;
						case 16:
							goto IL_03d9;
						case 17:
							goto IL_03df;
						case 18:
							goto IL_03e5;
						case 19:
							goto IL_03eb;
						case 20:
							goto IL_03fb;
						case 22:
							goto IL_0405;
						case 23:
							goto IL_0410;
						case 24:
							goto IL_041d;
						case 25:
							goto IL_0420;
						case 26:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 21:
						case 27:
							goto end_IL_0000;
						}
						goto default;
					}
					IL_0420:
					num2 = 25;
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(52701));
					break;
					IL_03fb:
					num2 = 20;
					Helpers.NoChartMessage();
					goto end_IL_0000;
					IL_0405:
					num2 = 22;
					wpfChartSize2 = new wpfChartSize(chart);
					goto IL_0410;
					IL_0410:
					num2 = 23;
					wpfChartSize2.ShowDialog();
					goto IL_041d;
					IL_0021:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0028;
					IL_0028:
					num2 = 4;
					chart = Helpers.SelectedChart();
					goto IL_0032;
					IL_0032:
					num2 = 5;
					if (chart == null)
					{
						goto IL_003a;
					}
					goto IL_03eb;
					IL_003a:
					num2 = 6;
					application = MH.A.Application;
					goto IL_0048;
					IL_0048:
					num2 = 7;
					worksheet = (Worksheet)application.ActiveSheet;
					goto IL_005a;
					IL_005a:
					num2 = 8;
					if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(worksheet.ChartObjects(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(52690), new object[0], null, null, null), 1, TextCompare: false))
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
						goto IL_00a9;
					}
					goto IL_03df;
					IL_041d:
					wpfChartSize2 = null;
					goto IL_0420;
					IL_00a9:
					num2 = 9;
					chartObject = (ChartObject)worksheet.ChartObjects(1);
					goto IL_00c2;
					IL_00c2:
					num2 = 10;
					pane = application.ActiveWindow.ActivePane;
					goto IL_00d7;
					IL_00d7:
					num2 = 11;
					if ((application.Intersect(chartObject.TopLeftCell, pane.VisibleRange, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null) & (application.Intersect(chartObject.BottomRightCell, pane.VisibleRange, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null))
					{
						goto IL_039f;
					}
					goto IL_03d6;
					IL_039f:
					num2 = 12;
					chart = chartObject.Chart;
					goto IL_03ac;
					IL_03ac:
					num2 = 13;
					chartObject.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
					goto IL_03c5;
					IL_03c5:
					num2 = 14;
					chart.ChartArea.Select();
					goto IL_03d6;
					IL_03d6:
					pane = null;
					goto IL_03d9;
					IL_03d9:
					num2 = 16;
					chartObject = null;
					goto IL_03df;
					IL_03df:
					num2 = 17;
					worksheet = null;
					goto IL_03e5;
					IL_03e5:
					num2 = 18;
					application = null;
					goto IL_03eb;
					IL_03eb:
					num2 = 19;
					if (chart == null)
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
						goto IL_03fb;
					}
					goto IL_0405;
					end_IL_0000_3:
					break;
				}
				num2 = 26;
				chart = null;
				break;
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 1212;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000:
			break;
		}
		if (num == 0)
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
			ProjectData.ClearProjectError();
			return;
		}
	}
}
