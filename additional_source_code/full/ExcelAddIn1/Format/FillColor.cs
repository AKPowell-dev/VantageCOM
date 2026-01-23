using System;
using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class FillColor
{
	public static void Cycle()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Application application = MH.A.Application;
		bool flag = false;
		if (!(application.Selection is Range))
		{
			return;
		}
		ColorCycle fillColorCycle = KH.A.FillColorCycle;
		int count = fillColorCycle.Colors.Count;
		checked
		{
			if (count > 0)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Range range = (Range)application.Selection;
				if (!Base.IsWorksheetProtected(range.Worksheet))
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
					application.ScreenUpdating = false;
					try
					{
						if (KH.A.UndoFill)
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
							flag = JH.A(range);
						}
						if (fillColorCycle.Colors[fillColorCycle.Index].RGB.Length > 0)
						{
							ColorCycle.Color color = fillColorCycle.Colors[fillColorCycle.Index];
							range.Interior.Pattern = color.Pattern;
							range.Interior.PatternColor = color.PatternOLE;
							range.Interior.Color = color.OLE;
							color = default(ColorCycle.Color);
						}
						else
						{
							range.Interior.ColorIndex = Constants.xlNone;
							range.Interior.Pattern = XlPattern.xlPatternNone;
						}
						if (flag)
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
							JH.A(range, VH.A(148068));
						}
						if (fillColorCycle.Index == 0)
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
							Base.LogActivity(fillColorCycle.Activity);
						}
						if (fillColorCycle.Index < count - 1)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								fillColorCycle.Index++;
								break;
							}
						}
						else
						{
							fillColorCycle.Index = 0;
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						Base.HandleFormattingException(ex2);
						ProjectData.ClearProjectError();
					}
					application.ScreenUpdating = true;
				}
				range = null;
			}
			fillColorCycle = null;
		}
	}

	internal static void A(int A)
	{
		try
		{
			FillColor.A(clsColors.ColorPalette[A].RGB);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Base.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	internal static void A(string A)
	{
		try
		{
			Color color = clsColors.RGB2Color(A);
			FillColor.A(ColorTranslator.ToOle(color), B: false);
			K.Settings.LastFillColor = color;
			KH.A.InvalidateControl(clsColors.LAST_FILL_COLOR_BUTTON);
			Base.LogActivity(VH.A(149603));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Base.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	internal static void A()
	{
		bool b = K.Settings.LastFillColor == Color.Transparent;
		try
		{
			A(ColorTranslator.ToOle(K.Settings.LastFillColor), b);
			Base.LogActivity(VH.A(149640));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Base.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	public static void None()
	{
		try
		{
			A(0, B: true);
			K.Settings.LastFillColor = Color.Transparent;
			KH.A.InvalidateControl(clsColors.LAST_FILL_COLOR_BUTTON);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Base.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	private static void A(int A, bool B)
	{
		object objectValue = RuntimeHelpers.GetObjectValue(MH.A.Application.Selection);
		if (objectValue is Range)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Range range = JH.A((Range)null);
			if (!Base.IsWorksheetProtected(range.Worksheet))
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
				bool flag = false;
				if (KH.A.UndoFill)
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
					flag = JH.A(range);
				}
				if (!B)
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
					range.Interior.Color = A;
				}
				else
				{
					range.Interior.ColorIndex = Constants.xlNone;
				}
				if (flag)
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
					JH.A(range, VH.A(148068));
				}
			}
			range = null;
		}
		else if (objectValue is Series)
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
			FillColor.A(((Series)objectValue).Format, A, B);
		}
		else if (objectValue is Microsoft.Office.Interop.Excel.Point)
		{
			FillColor.A(((Microsoft.Office.Interop.Excel.Point)objectValue).Format, A, B);
		}
		else if (objectValue is ChartArea)
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
			FillColor.A(((ChartArea)objectValue).Format, A, B);
		}
		else if (objectValue is PlotArea)
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
			FillColor.A(((PlotArea)objectValue).Format, A, B);
		}
		else if (objectValue is DataLabels)
		{
			FillColor.A(((DataLabels)objectValue).Format, A, B);
		}
		else if (objectValue is DataLabel)
		{
			FillColor.A(((DataLabel)objectValue).Format, A, B);
		}
		else if (objectValue is Axis)
		{
			FillColor.A(((Axis)objectValue).Format, A, B);
		}
		else if (objectValue is AxisTitle)
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
			FillColor.A(((AxisTitle)objectValue).Format, A, B);
		}
		else if (objectValue is ChartTitle)
		{
			FillColor.A(((ChartTitle)objectValue).Format, A, B);
		}
		else if (objectValue is Legend)
		{
			FillColor.A(((Legend)objectValue).Format, A, B);
		}
		else if (!(objectValue is LegendEntry) && objectValue is DataTable)
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
			FillColor.A(((DataTable)objectValue).Format, A, B);
		}
		objectValue = null;
	}

	private static void A(ChartFormat A, int B, bool C)
	{
		Microsoft.Office.Interop.Excel.FillFormat fill = A.Fill;
		if (!C)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			fill.Visible = MsoTriState.msoTrue;
			fill.ForeColor.RGB = B;
			fill.BackColor.RGB = B;
		}
		else
		{
			fill.Visible = MsoTriState.msoFalse;
		}
		fill = null;
	}

	public static bool HasFill(Interior interior)
	{
		return Operators.ConditionalCompareObjectNotEqual(interior.ColorIndex, Constants.xlNone, TextCompare: false);
	}
}
