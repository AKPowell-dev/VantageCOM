using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Config;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class Borders
{
	[CompilerGenerated]
	private static int m_A;

	[CompilerGenerated]
	private static int m_B;

	[CompilerGenerated]
	private static int C;

	[CompilerGenerated]
	private static int D;

	internal static int BorderIndex
	{
		[CompilerGenerated]
		get
		{
			return Borders.m_A;
		}
		[CompilerGenerated]
		set
		{
			Borders.m_A = value;
		}
	}

	internal static int CycleIndex
	{
		[CompilerGenerated]
		get
		{
			return Borders.m_B;
		}
		[CompilerGenerated]
		set
		{
			Borders.m_B = value;
		}
	}

	internal static int OutsideBorderCycleIndex
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	internal static int InsideBorderCycleIndex
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[CompilerGenerated]
		set
		{
			D = value;
		}
	}

	public static void CycleTop()
	{
		A(XlBordersIndex.xlEdgeTop);
	}

	public static void CycleBottom()
	{
		A(XlBordersIndex.xlEdgeBottom);
	}

	public static void CycleLeft()
	{
		A(XlBordersIndex.xlEdgeLeft);
	}

	public static void CycleRight()
	{
		A(XlBordersIndex.xlEdgeRight);
	}

	private static void A(XlBordersIndex A)
	{
		Application application = MH.A.Application;
		Range range3 = default(Range);
		Range range;
		try
		{
			application.ScreenUpdating = false;
			range = JH.A((Range)null);
			bool flag = JH.A(range);
			if (A != (XlBordersIndex)BorderIndex)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				CycleIndex = 0;
			}
			string[] array = Strings.Split(KH.A.BorderStyleCycle[CycleIndex], VH.A(2378));
			BorderIndex = (int)A;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = range.Areas.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range2 = (Range)enumerator.Current;
					switch (A)
					{
					case XlBordersIndex.xlEdgeTop:
						range3 = (Range)NewLateBinding.LateGet(range2.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(62391), new object[0], null, null, null);
						range2.get_Offset((object)(-1), (object)0).Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
						break;
					case XlBordersIndex.xlEdgeBottom:
						range3 = (Range)NewLateBinding.LateGet(range2.Rows[RuntimeHelpers.GetObjectValue(range2.Rows.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(62391), new object[0], null, null, null);
						range2.get_Offset((object)1, (object)0).Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
						break;
					case XlBordersIndex.xlEdgeLeft:
						range3 = (Range)NewLateBinding.LateGet(range2.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(62391), new object[0], null, null, null);
						range2.get_Offset((object)0, (object)(-1)).Borders[XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
						break;
					case XlBordersIndex.xlEdgeRight:
						range3 = (Range)NewLateBinding.LateGet(range2.Columns[RuntimeHelpers.GetObjectValue(range2.Columns.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(62391), new object[0], null, null, null);
						range2.get_Offset((object)0, (object)1).Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
						break;
					}
					range2 = null;
					Border border = range3.Borders[A];
					border.LineStyle = Borders.GetLineStyle(array[0]);
					if (Conversions.ToDouble(array[0]) != -4142.0)
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
						border.Weight = array[1];
						border.Color = KH.A.DefaultBorderColor;
					}
					border = null;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_033d;
					}
					continue;
					end_IL_033d:
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			if (flag)
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
				if (KH.A.UndoBorders)
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
					JH.A(range, VH.A(146542));
				}
			}
			checked
			{
				if (CycleIndex == KH.A.BorderStyleCycle.Count - 1)
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
					CycleIndex = 0;
				}
				else
				{
					CycleIndex++;
				}
				if (CycleIndex == 1)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						Base.LogActivity(VH.A(146557));
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application.ScreenUpdating = true;
		application = null;
		range3 = null;
		range = null;
	}

	public static void Outside()
	{
		Application application = MH.A.Application;
		checked
		{
			Range range;
			List<XlBordersIndex> list;
			try
			{
				string[] array = Strings.Split(KH.A.BorderStyleCycle[OutsideBorderCycleIndex], VH.A(2378));
				range = (Range)application.Selection;
				application.ScreenUpdating = false;
				bool flag = JH.A(range);
				list = new List<XlBordersIndex>
				{
					XlBordersIndex.xlEdgeTop,
					XlBordersIndex.xlEdgeBottom,
					XlBordersIndex.xlEdgeLeft,
					XlBordersIndex.xlEdgeRight
				};
				IEnumerator enumerator = default(IEnumerator);
				Range range4 = default(Range);
				try
				{
					enumerator = range.Areas.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range range2 = (Range)enumerator.Current;
						foreach (XlBordersIndex item in list)
						{
							Range range3 = range2;
							switch (item)
							{
							case XlBordersIndex.xlEdgeTop:
								range4 = (Range)NewLateBinding.LateGet(range3.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(62391), new object[0], null, null, null);
								break;
							case XlBordersIndex.xlEdgeBottom:
								range4 = (Range)NewLateBinding.LateGet(range3.Rows[RuntimeHelpers.GetObjectValue(range3.Rows.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(62391), new object[0], null, null, null);
								break;
							case XlBordersIndex.xlEdgeLeft:
								range4 = (Range)NewLateBinding.LateGet(range3.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(62391), new object[0], null, null, null);
								break;
							case XlBordersIndex.xlEdgeRight:
								range4 = (Range)NewLateBinding.LateGet(range3.Columns[RuntimeHelpers.GetObjectValue(range3.Columns.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(62391), new object[0], null, null, null);
								break;
							}
							range3 = null;
							Border border = range4.Borders[item];
							border.LineStyle = Borders.GetLineStyle(array[0]);
							if (Conversions.ToDouble(array[0]) != -4142.0)
							{
								border.Weight = array[1];
								border.Color = KH.A.DefaultBorderColor;
							}
							border = null;
						}
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
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				range4 = null;
				if (flag)
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
					if (KH.A.UndoBorders)
					{
						JH.A(range, VH.A(146542));
					}
				}
				if (OutsideBorderCycleIndex == KH.A.BorderStyleCycle.Count - 1)
				{
					OutsideBorderCycleIndex = 0;
				}
				else
				{
					OutsideBorderCycleIndex++;
				}
				if (OutsideBorderCycleIndex == 1)
				{
					Base.LogActivity(VH.A(146557));
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			application.ScreenUpdating = true;
			application = null;
			range = null;
			list = null;
		}
	}

	public static void Inside()
	{
		Application application = MH.A.Application;
		checked
		{
			Range range;
			List<XlBordersIndex> list;
			try
			{
				string[] array = Strings.Split(KH.A.BorderStyleCycle[InsideBorderCycleIndex], VH.A(2378));
				range = (Range)application.Selection;
				application.ScreenUpdating = false;
				bool flag = JH.A(range);
				list = new List<XlBordersIndex>
				{
					XlBordersIndex.xlInsideHorizontal,
					XlBordersIndex.xlInsideVertical
				};
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = range.Areas.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range range2 = (Range)enumerator.Current;
						if (!Operators.ConditionalCompareObjectGreater(range2.CountLarge, 1, TextCompare: false))
						{
							continue;
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						using List<XlBordersIndex>.Enumerator enumerator2 = list.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							XlBordersIndex current = enumerator2.Current;
							Border border = range2.Borders[current];
							border.LineStyle = Borders.GetLineStyle(array[0]);
							if (Conversions.ToDouble(array[0]) != -4142.0)
							{
								border.Weight = array[1];
								border.Color = KH.A.DefaultBorderColor;
							}
							border = null;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_0145;
							}
							continue;
							end_IL_0145:
							break;
						}
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_016d;
						}
						continue;
						end_IL_016d:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				if (flag)
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
					if (KH.A.UndoBorders)
					{
						JH.A(range, VH.A(146542));
					}
				}
				if (InsideBorderCycleIndex == KH.A.BorderStyleCycle.Count - 1)
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
					InsideBorderCycleIndex = 0;
				}
				else
				{
					InsideBorderCycleIndex++;
				}
				if (InsideBorderCycleIndex == 1)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						Base.LogActivity(VH.A(146557));
						break;
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			application.ScreenUpdating = true;
			application = null;
			range = null;
			list = null;
		}
	}

	public static void None()
	{
		try
		{
			Range range = JH.A((Range)null);
			bool num = JH.A(range);
			range.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
			if (num && KH.A.UndoBorders)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				JH.A(range, VH.A(146542));
			}
			range = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void CycleColor()
	{
		if (!Licensing.AllowRestrictedMode())
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
				if (application.Selection is Range)
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
					ColorCycle borderColorCycle = KH.A.BorderColorCycle;
					int count = borderColorCycle.Colors.Count;
					if (count > 0)
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
						application.ScreenUpdating = false;
						try
						{
							B(borderColorCycle.Colors[borderColorCycle.Index].OLE);
							if (borderColorCycle.Index == 0)
							{
								Base.LogActivity(borderColorCycle.Activity);
							}
							if (borderColorCycle.Index < count - 1)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									borderColorCycle.Index++;
									break;
								}
							}
							else
							{
								borderColorCycle.Index = 0;
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						application.ScreenUpdating = true;
					}
					borderColorCycle = null;
				}
				application = null;
				return;
			}
		}
	}

	internal static void A(int A)
	{
		try
		{
			Borders.A(clsColors.ColorPalette[A].RGB);
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
			Borders.A(ColorTranslator.ToOle(color), B: false);
			K.Settings.LastBorderColor = color;
			KH.A.InvalidateControl(clsColors.LAST_BORDER_COLOR_BUTTON);
			Base.LogActivity(VH.A(146594));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	internal static void A()
	{
		bool b = K.Settings.LastBorderColor == Color.Transparent;
		try
		{
			A(ColorTranslator.ToOle(K.Settings.LastBorderColor), b);
			Base.LogActivity(VH.A(146635));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	public static void NoBorder()
	{
		try
		{
			A(0, B: true);
			K.Settings.LastBorderColor = Color.Transparent;
			KH.A.InvalidateControl(clsColors.LAST_BORDER_COLOR_BUTTON);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	private static void A(int A, bool B)
	{
		Application application = MH.A.Application;
		object objectValue = RuntimeHelpers.GetObjectValue(application.Selection);
		if (objectValue is Range)
		{
			application.ScreenUpdating = false;
			if (!B)
			{
				try
				{
					Borders.B(A);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					clsReporting.LogException(ex2);
					ProjectData.ClearProjectError();
				}
			}
			else
			{
				None();
			}
			application.ScreenUpdating = true;
		}
		else if (objectValue is Series)
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
			Borders.A(((Series)objectValue).Format, A, B);
		}
		else if (objectValue is ChartArea)
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
			Borders.A(((ChartArea)objectValue).Format, A, B);
		}
		else if (objectValue is PlotArea)
		{
			Borders.A(((PlotArea)objectValue).Format, A, B);
		}
		else if (objectValue is Gridlines)
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
			Borders.A(((Gridlines)objectValue).Format, A, B);
		}
		else if (objectValue is DataLabels)
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
			Borders.A(((DataLabels)objectValue).Format, A, B);
		}
		else if (objectValue is DataLabel)
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
			Borders.A(((DataLabel)objectValue).Format, A, B);
		}
		else if (!(objectValue is Axis))
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
			if (objectValue is AxisTitle)
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
				Borders.A(((AxisTitle)objectValue).Format, A, B);
			}
			else if (objectValue is ChartTitle)
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
				Borders.A(((ChartTitle)objectValue).Format, A, B);
			}
			else if (objectValue is Legend)
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
				Legend legend = (Legend)objectValue;
				if (!B)
				{
					Borders.A(legend.Format, A, B);
				}
				else
				{
					legend.Format.Line.Transparency = 1f;
				}
				legend = null;
			}
			else if (!(objectValue is LegendEntry))
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
				if (objectValue is DataTable)
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
					Borders.A(((DataTable)objectValue).Format, A, B);
				}
			}
		}
		objectValue = null;
		application = null;
	}

	private static void A(ChartFormat A, int B, bool C)
	{
		LineFormat line = A.Line;
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
			line.Visible = MsoTriState.msoTrue;
			line.Transparency = 0f;
			line.ForeColor.RGB = B;
			line.BackColor.RGB = B;
		}
		else
		{
			line.Visible = MsoTriState.msoFalse;
		}
		line = null;
	}

	private static void B(int A)
	{
		Range range = JH.A((Range)null);
		bool flag = JH.A(range);
		List<XlBordersIndex> list = new List<XlBordersIndex>
		{
			XlBordersIndex.xlEdgeTop,
			XlBordersIndex.xlEdgeBottom,
			XlBordersIndex.xlEdgeLeft,
			XlBordersIndex.xlEdgeRight
		};
		IEnumerator enumerator = range.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Range range2 = (Range)enumerator.Current;
				using List<XlBordersIndex>.Enumerator enumerator2 = list.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					XlBordersIndex current = enumerator2.Current;
					Border border = range2.Borders[current];
					if (Operators.ConditionalCompareObjectNotEqual(border.LineStyle, Microsoft.Office.Interop.Excel.Constants.xlNone, TextCompare: false))
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
						border.Color = A;
					}
					border = null;
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_00c5;
					}
					continue;
					end_IL_00c5:
					break;
				}
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_00ea;
				}
				continue;
				end_IL_00ea:
				break;
			}
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
		if (flag && KH.A.UndoBorders)
		{
			JH.A(range, VH.A(146542));
		}
		range = null;
		list = null;
	}

	public static bool HasTopBorder(Microsoft.Office.Interop.Excel.Borders borders)
	{
		return A(borders, XlBordersIndex.xlEdgeTop);
	}

	public static bool HasBottomBorder(Microsoft.Office.Interop.Excel.Borders borders)
	{
		return A(borders, XlBordersIndex.xlEdgeBottom);
	}

	public static bool HasLeftBorder(Microsoft.Office.Interop.Excel.Borders borders)
	{
		return A(borders, XlBordersIndex.xlEdgeLeft);
	}

	public static bool HasRightBorder(Microsoft.Office.Interop.Excel.Borders borders)
	{
		return A(borders, XlBordersIndex.xlEdgeRight);
	}

	private static bool A(Microsoft.Office.Interop.Excel.Borders A, XlBordersIndex B)
	{
		return HasBorder(A[B]);
	}

	public static bool HasBorder(Border border)
	{
		return Operators.ConditionalCompareObjectNotEqual(border.LineStyle, Microsoft.Office.Interop.Excel.Constants.xlNone, TextCompare: false);
	}
}
