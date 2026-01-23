using System;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1;

public sealed class clsCharts
{
	public static bool UsesPointsForSeriesClrs(Chart cht)
	{
		return A(cht, new int[9] { 5, -4102, 68, 71, -4120, 69, 70, 80, 119 });
	}

	public static bool UsesLegendsForSeriesClrs(Chart cht)
	{
		if (!UsesLegendLinesForSeriesClrs(cht))
		{
			return UsesLegendFillsForSeriesClrs(cht);
		}
		return true;
	}

	public static bool UsesFormatFillForSeriesClrs(Chart cht)
	{
		if (!UsesLegendsForSeriesClrs(cht))
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
					return !UsesPointsForSeriesClrs(cht);
				}
			}
		}
		return false;
	}

	public static bool UsesLegendFillsForSeriesClrs(Chart cht)
	{
		return A(cht, new int[3] { 83, 85, 82 });
	}

	public static bool UsesLegendLinesForSeriesClrs(Chart cht)
	{
		return A(cht, new int[10] { 84, 86, -4151, 81, 4, 63, 64, 65, 66, 67 });
	}

	public static bool UsesMarkers(Chart cht)
	{
		return A(cht, new int[4] { 81, 65, 66, 67 });
	}

	public static bool SeriesClrsAreUnusable(Chart cht)
	{
		return A(cht, new int[1] { 119 });
	}

	public static bool CanIgnoreLegendClrs(Chart cht)
	{
		if (cht.HasLegend)
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
					return UsesPointsForSeriesClrs(cht);
				}
			}
		}
		return true;
	}

	internal static bool A(Chart A, int[] B)
	{
		return B.Contains((int)A.ChartType);
	}

	internal static bool A(Shape A, int[] B)
	{
		bool result;
		try
		{
			result = !clsCharts.A(A.Chart, B);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result = true;
			ProjectData.ClearProjectError();
		}
		finally
		{
		}
		return result;
	}
}
