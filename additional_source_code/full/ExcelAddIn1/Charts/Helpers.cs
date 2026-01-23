using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Auth;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class Helpers
{
	public static Chart SelectedChart()
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Chart activeChart = default(Chart);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					break;
				case 59:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 3:
							goto end_IL_0000_3;
						}
						goto default;
					}
					end_IL_0000_2:
					break;
				}
				num2 = 2;
				activeChart = MH.A.Application.ActiveChart;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 59;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
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
			ProjectData.ClearProjectError();
		}
		return activeChart;
	}

	public static (bool AreValid, List<Chart> ChartList) SelectedCharts(Func<Chart, bool> isValidFunc)
	{
		(bool, List<Chart>) tuple = (true, new List<Chart>());
		(bool, List<Chart>) result;
		try
		{
			object objectValue = RuntimeHelpers.GetObjectValue(MH.A.Application.Selection);
			if (Operators.CompareString(Versioned.TypeName(RuntimeHelpers.GetObjectValue(objectValue)), VH.A(56245), TextCompare: false) == 0)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					ShapeRange shapeRange = (ShapeRange)NewLateBinding.LateGet(objectValue, null, VH.A(56274), new object[0], null, null, null);
					if (shapeRange.Count <= 0)
					{
						break;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						foreach (Shape item in shapeRange)
						{
							if (item.HasChart != MsoTriState.msoTrue)
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
								break;
							}
							if (isValidFunc != null && !isValidFunc(item.Chart))
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									shapeRange = null;
									tuple.Item1 = false;
									result = tuple;
									break;
								}
								goto end_IL_0091;
							}
							tuple.Item2.Add(item.Chart);
						}
						goto end_IL_004e;
						continue;
						end_IL_0091:
						break;
					}
					goto IL_01a7;
					continue;
					end_IL_004e:
					break;
				}
			}
			else
			{
				Chart chart = SelectedChart();
				if (chart == null)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						tuple.Item2 = null;
						result = tuple;
						break;
					}
					goto IL_01a7;
				}
				if (isValidFunc != null)
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
					if (!isValidFunc(chart))
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							chart = null;
							result = tuple;
							break;
						}
						goto IL_01a7;
					}
				}
				tuple.Item2.Add(chart);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			tuple.Item2.Clear();
			result = tuple;
			ProjectData.ClearProjectError();
			goto IL_01a7;
		}
		result = tuple;
		goto IL_01a7;
		IL_01a7:
		return result;
	}

	internal static string[] A(Series A)
	{
		return Helpers.A(A.Formula);
	}

	internal static string[] A(string A)
	{
		return Strings.Split(A.Replace(VH.A(78994), ""), VH.A(2378));
	}

	internal static void A()
	{
		Forms.WarningMessage(VH.A(79011));
	}

	public static void NoChartMessage()
	{
		Forms.WarningMessage(VH.A(56295));
	}

	internal static void B()
	{
		Forms.ErrorMessage(Charts.NotImplementedMsgText());
	}

	internal static void A(Application A, Chart B)
	{
		A.ScreenUpdating = false;
		if (Helpers.A(A))
		{
			try
			{
				ChartObject obj = (ChartObject)B.Parent;
				obj.Width += 1.0;
				ChartObject chartObject = obj;
				obj.Width = chartObject.Width - 1.0;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		A.ScreenUpdating = true;
	}

	private static bool A(Application A)
	{
		return true;
	}

	internal static bool A()
	{
		return Access.AllowExcelOperation((PlanType)4, (Restriction)1, false);
	}
}
