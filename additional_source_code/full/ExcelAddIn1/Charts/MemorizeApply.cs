using System;
using System.Collections;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class MemorizeApply
{
	public enum MemorizedProperty
	{
		ChartSize = 1,
		PlotSize,
		PlotPosition,
		All
	}

	private struct FD
	{
		public double A;

		public double B;
	}

	private struct GD
	{
		public double A;

		public double B;
	}

	[CompilerGenerated]
	private static FD m_A;

	[CompilerGenerated]
	private static FD m_B;

	[CompilerGenerated]
	private static GD m_A;

	private static FD MemorizedChartSize
	{
		[CompilerGenerated]
		get
		{
			return MemorizeApply.m_A;
		}
		[CompilerGenerated]
		set
		{
			MemorizeApply.m_A = value;
		}
	} = default(FD);

	private static FD MemorizedPlotSize
	{
		[CompilerGenerated]
		get
		{
			return MemorizeApply.m_B;
		}
		[CompilerGenerated]
		set
		{
			MemorizeApply.m_B = value;
		}
	} = default(FD);

	private static GD MemorizedPlotPosition
	{
		[CompilerGenerated]
		get
		{
			return MemorizeApply.m_A;
		}
		[CompilerGenerated]
		set
		{
			MemorizeApply.m_A = value;
		}
	} = default(GD);

	public static void Memorize()
	{
		Chart chart = null;
		try
		{
			chart = Helpers.SelectedChart();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (chart != null)
		{
			FD fD = default(FD);
			GD memorizedPlotPosition = default(GD);
			Chart chart2 = chart;
			fD.A = chart2.ChartArea.Width;
			fD.B = chart2.ChartArea.Height;
			MemorizedChartSize = fD;
			fD.A = chart2.PlotArea.InsideWidth;
			fD.B = chart2.PlotArea.InsideHeight;
			MemorizedPlotSize = fD;
			memorizedPlotPosition.A = chart2.PlotArea.InsideLeft;
			memorizedPlotPosition.B = chart2.PlotArea.InsideTop;
			MemorizedPlotPosition = memorizedPlotPosition;
			chart2 = null;
			chart = null;
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(56155));
		}
	}

	public static void SetToMemorized(MemorizedProperty mcp)
	{
		Chart chart = null;
		if (MemorizedChartSize.A > 0.0)
		{
			try
			{
				chart = Helpers.SelectedChart();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (chart != null)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (chart.Application.ActiveSheet is Chart)
						{
							Helpers.A();
						}
						else
						{
							try
							{
								A(chart, mcp);
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								Forms.ErrorMessage(ex4.Message);
								clsReporting.LogException(ex4);
								ProjectData.ClearProjectError();
							}
						}
						chart = null;
						clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(56184));
						return;
					}
				}
			}
			IEnumerator enumerator = default(IEnumerator);
			if (Operators.CompareString(Versioned.TypeName(RuntimeHelpers.GetObjectValue(MH.A.Application.Selection)), VH.A(56245), TextCompare: false) == 0)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						try
						{
							try
							{
								enumerator = ((IEnumerable)NewLateBinding.LateGet(MH.A.Application.Selection, null, VH.A(56274), new object[0], null, null, null)).GetEnumerator();
								while (enumerator.MoveNext())
								{
									Shape shape = (Shape)enumerator.Current;
									if (shape.HasChart == MsoTriState.msoTrue)
									{
										A(shape.Chart, mcp);
									}
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_014f;
									}
									continue;
									end_IL_014f:
									break;
								}
							}
							finally
							{
								if (enumerator is IDisposable)
								{
									while (true)
									{
										switch (1)
										{
										case 0:
											break;
										default:
											(enumerator as IDisposable).Dispose();
											goto end_IL_0163;
										}
										continue;
										end_IL_0163:
										break;
									}
								}
							}
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
							Forms.ErrorMessage(ex6.Message);
							clsReporting.LogException(ex6);
							ProjectData.ClearProjectError();
						}
						clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(56184));
						return;
					}
				}
			}
			Forms.WarningMessage(VH.A(56295));
		}
		else
		{
			Forms.WarningMessage(VH.A(56332));
		}
	}

	private static void A(Chart A, MemorizedProperty B)
	{
		switch (B)
		{
		case MemorizedProperty.ChartSize:
			MemorizeApply.A(A);
			break;
		case MemorizedProperty.PlotSize:
			MemorizeApply.B(A);
			break;
		case MemorizedProperty.PlotPosition:
			C(A);
			break;
		case MemorizedProperty.All:
			MemorizeApply.A(A);
			C(A);
			MemorizeApply.B(A);
			C(A);
			break;
		}
	}

	private static void A(Chart A)
	{
		Charts.ChangeChartAreaSize(A, (double?)MemorizedChartSize.A, (double?)MemorizedChartSize.B);
	}

	private static void B(Chart A)
	{
		PlotArea plotArea = A.PlotArea;
		plotArea.InsideWidth = MemorizedPlotSize.A;
		plotArea.InsideHeight = MemorizedPlotSize.B;
		_ = null;
	}

	private static void C(Chart A)
	{
		PlotArea plotArea = A.PlotArea;
		plotArea.InsideLeft = MemorizedPlotPosition.A;
		plotArea.InsideTop = MemorizedPlotPosition.B;
		_ = null;
	}
}
