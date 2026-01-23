using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Xml;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format.ContextualPopups;

public sealed class Base
{
	private static clsDisplay m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("ThisChart")]
	private static Chart m_A;

	private static readonly int m_A = 88;

	private static readonly int m_B = 90;

	public static clsDisplay Dpi
	{
		get
		{
			return Base.m_A;
		}
		set
		{
			Base.m_A = value;
		}
	}

	private static Chart ThisChart
	{
		[CompilerGenerated]
		get
		{
			return Base.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			Base.m_A = value;
		}
	}

	public static void SheetActivate(object Sh)
	{
		if (!(Sh is Worksheet))
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
			RemoveChartHandlers((Worksheet)Sh);
			AddChartHandlers((Worksheet)Sh);
			return;
		}
	}

	public static void SheetSelectionChange(object Sh, Range Target)
	{
		if (!(Sh is Worksheet))
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
			RemoveChartHandlers((Worksheet)Sh);
			AddChartHandlers((Worksheet)Sh);
			return;
		}
	}

	public static void AddChartHandlers(Worksheet ws)
	{
		int num = Conversions.ToInteger(NewLateBinding.LateGet(ws.ChartObjects(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(52690), new object[0], null, null, null));
		for (int i = 1; i <= num; i = checked(i + 1))
		{
			new ComAwareEventInfo(typeof(ChartEvents_Event), VH.A(144800)).AddEventHandler(((ChartObject)ws.ChartObjects(i)).Chart, new ChartEvents_MouseUpEventHandler(ChartMouseUp));
		}
	}

	public static void RemoveChartHandlers(Worksheet ws)
	{
		int num = Conversions.ToInteger(NewLateBinding.LateGet(ws.ChartObjects(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(52690), new object[0], null, null, null));
		for (int i = 1; i <= num; i = checked(i + 1))
		{
			new ComAwareEventInfo(typeof(ChartEvents_Event), VH.A(144800)).RemoveEventHandler(((ChartObject)ws.ChartObjects(i)).Chart, new ChartEvents_MouseUpEventHandler(ChartMouseUp));
		}
	}

	private static void A(long A, long B, long C)
	{
		switch ((XlChartItem)checked((int)A))
		{
		default:
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
			break;
		case XlChartItem.xlSeries:
			if (C > 0)
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
				Interaction.MsgBox(VH.A(144815));
			}
			else
			{
				Interaction.MsgBox(VH.A(144826));
			}
			break;
		case XlChartItem.xlChartArea:
		case XlChartItem.xlChartTitle:
		case XlChartItem.xlPlotArea:
			break;
		}
		Interaction.MsgBox(((XlChartItem)checked((int)A)/*cast due to .constrained prefix*/).ToString());
	}

	private static void A()
	{
		try
		{
			_ = MH.A.Application.ActiveChart;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void ChartMouseUp(int Button, int Shift, int x, int y)
	{
		//IL_002a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0034: Expected O, but got Unknown
		object objectValue = RuntimeHelpers.GetObjectValue(MH.A.Application.Selection);
		XmlDocument settingsXml = KH.A.SettingsXml;
		Dpi = new clsDisplay();
		try
		{
			if (objectValue is PlotArea)
			{
				double C = default(double);
				double D = default(double);
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
					PlotArea plotArea = (PlotArea)objectValue;
					ChartObject a = (ChartObject)NewLateBinding.LateGet(plotArea.Parent, null, VH.A(8701), new object[0], null, null, null);
					float num = A(a);
					float num2 = B(a);
					B((float)(plotArea.InsideLeft + (double)num), (float)(plotArea.InsideTop + (double)num2), ref C, ref D);
					Rectangle rectangle = A(a);
					C = rectangle.Left;
					D = rectangle.Top;
					XmlNode nd = settingsXml.DocumentElement.SelectSingleNode(VH.A(144839));
					wpfPlotArea obj = new wpfPlotArea(nd, ((PlotArea)objectValue).Format.Fill.ForeColor.RGB, plotArea);
					obj.Top = D + (double)y;
					obj.Left = C + (double)x;
					obj.Show();
					_ = null;
					plotArea = null;
					break;
				}
			}
			else if (!(objectValue is Series))
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					_ = objectValue is SeriesCollection;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Interaction.MsgBox(ex2.Message);
			ProjectData.ClearProjectError();
		}
		objectValue = null;
		settingsXml = null;
	}

	private static Rect A(float A, float B, float C, float D)
	{
		Microsoft.Office.Interop.Excel.Window activeWindow = MH.A.Application.ActiveWindow;
		checked
		{
			int num = (int)Math.Round((double)activeWindow.PointsToScreenPixelsX((int)Math.Round(A)) / Dpi.X);
			int num2 = (int)Math.Round((double)activeWindow.PointsToScreenPixelsY((int)Math.Round(B)) / Dpi.Y);
			Rect result = new Rect(num, num2, (double)activeWindow.PointsToScreenPixelsX((int)Math.Round(A + C)) / Dpi.X - (double)num, (double)activeWindow.PointsToScreenPixelsY((int)Math.Round(B + D)) / Dpi.Y - (double)num2);
			activeWindow = null;
			return result;
		}
	}

	private static void B(float A, float B, ref double C, ref double D)
	{
		Microsoft.Office.Interop.Excel.Window activeWindow = MH.A.Application.ActiveWindow;
		checked
		{
			C = (double)activeWindow.PointsToScreenPixelsX((int)Math.Round(A)) / Dpi.X;
			D = (double)activeWindow.PointsToScreenPixelsY((int)Math.Round(B)) / Dpi.Y;
			activeWindow = null;
		}
	}

	private static float A(ChartObject A)
	{
		return (float)(A.Left + A.Chart.ChartArea.Left);
	}

	private static float B(ChartObject A)
	{
		return (float)(A.Top + A.Chart.ChartArea.Top);
	}

	[DllImport("gdi32.dll", EntryPoint = "GetDeviceCaps")]
	private static extern int A(IntPtr A, int B);

	[DllImport("user32.dll", EntryPoint = "GetDC")]
	private static extern IntPtr A(IntPtr A);

	[DllImport("user32.dll", EntryPoint = "ReleaseDC")]
	private static extern bool A(IntPtr A, IntPtr B);

	private static Rectangle A(ChartObject A)
	{
		Microsoft.Office.Interop.Excel.Application application = A.Application;
		Microsoft.Office.Interop.Excel.Window activeWindow = application.ActiveWindow;
		IntPtr intPtr = Base.A(IntPtr.Zero);
		long num = Base.A(intPtr, Base.m_A);
		long num2 = Base.A(intPtr, Base.m_B);
		Base.A(IntPtr.Zero, intPtr);
		double num3 = application.InchesToPoints(1.0);
		double num4 = Conversions.ToDouble(Operators.DivideObject(activeWindow.Zoom, 100));
		int num5 = activeWindow.PointsToScreenPixelsX(0);
		int num6 = activeWindow.PointsToScreenPixelsY(0);
		int x = Convert.ToInt32((double)num5 + A.Left * num4 * (double)num / num3);
		int width;
		try
		{
			width = 100;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			width = 10000;
			ProjectData.ClearProjectError();
		}
		int y = Convert.ToInt32((double)num6 + A.Top * num4 * (double)num2 / num3);
		int height;
		try
		{
			height = 100;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			height = 10000;
			ProjectData.ClearProjectError();
		}
		application = null;
		return new Rectangle(x, y, width, height);
	}
}
