using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Media;
using System.Xml;
using A;
using ExcelAddIn1.Format;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class TargetBand
{
	private enum PD
	{
		A,
		B
	}

	private struct QD
	{
		public int A;

		public int B;

		public float A;

		public float B;
	}

	public static void Add()
	{
		if (!Licensing.AllowChartAddOnOperation())
		{
			return;
		}
		Chart chart = Helpers.SelectedChart();
		string text;
		Series series;
		if (chart != null)
		{
			text = "";
			series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Item(1);
			XlChartType chartType = series.ChartType;
			if (chartType <= XlChartType.xlLine)
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
				if (chartType != XlChartType.xlXYScatter && chartType != XlChartType.xlLine)
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
					goto IL_010b;
				}
			}
			else if ((uint)(chartType - 51) > 1u)
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
				if ((uint)(chartType - 57) > 1u)
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
					switch (chartType)
					{
					case XlChartType.xlLineStacked:
					case XlChartType.xlLineMarkers:
					case XlChartType.xlLineMarkersStacked:
					case XlChartType.xlXYScatterSmooth:
					case XlChartType.xlXYScatterSmoothNoMarkers:
					case XlChartType.xlXYScatterLines:
					case XlChartType.xlXYScatterLinesNoMarkers:
						break;
					default:
						goto IL_010b;
					}
				}
			}
			if (series.AxisGroup == XlAxisGroup.xlSecondary)
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
				text = VH.A(68500);
			}
			goto IL_011a;
		}
		Helpers.NoChartMessage();
		return;
		IL_011a:
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = ((IEnumerable)chart.FullSeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
			while (true)
			{
				if (enumerator.MoveNext())
				{
					series = (Series)enumerator.Current;
					if (series.AxisGroup == XlAxisGroup.xlSecondary)
					{
						text = VH.A(68765);
						break;
					}
					continue;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_016f;
					}
					continue;
					end_IL_016f:
					break;
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
		if (text.Length > 0)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					Forms.WarningMessage(text);
					chart = null;
					series = null;
					return;
				}
			}
		}
		Application application = MH.A.Application;
		Range range = null;
		Range range2 = null;
		IEnumerator enumerator2 = default(IEnumerator);
		bool d = default(bool);
		try
		{
			enumerator2 = ((IEnumerable)chart.FullSeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
			while (enumerator2.MoveNext())
			{
				series = (Series)enumerator2.Current;
				string[] array = Helpers.A(series);
				Range range3;
				try
				{
					range3 = ((_Application)application).get_Range((object)array[2], RuntimeHelpers.GetObjectValue(Missing.Value));
					d = range3.Rows.Count == 1;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					range3 = null;
					ProjectData.ClearProjectError();
				}
				if (range == null)
				{
					range = range3;
				}
				else if (range3 != null)
				{
					range = application.Union(range, range3, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				Range range4;
				try
				{
					range4 = ((_Application)application).get_Range((object)array[0], RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					range4 = null;
					ProjectData.ClearProjectError();
				}
				if (range2 == null)
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
					range2 = range4;
				}
				else if (range4 != null)
				{
					range2 = application.Union(range2, range4, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				}
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_0560;
				}
				continue;
				end_IL_0560:
				break;
			}
		}
		finally
		{
			if (enumerator2 is IDisposable)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					(enumerator2 as IDisposable).Dispose();
					break;
				}
			}
		}
		bool B = false;
		QD e = A(chart, ref B);
		if (B)
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			bool autoExpandListRange = application.AutoCorrect.AutoExpandListRange;
			application.AutoCorrect.AutoExpandListRange = false;
			application.CutCopyMode = (XlCutCopyMode)0;
			application.ScreenUpdating = false;
			double num = -4142.0;
			if (Conversions.ToBoolean(((_Chart)chart).get_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlValue, (object)XlAxisGroup.xlPrimary)))
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
				Axis axis = (Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue);
				if (axis.MaximumScaleIsAuto)
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
					num = axis.MaximumScale;
				}
				axis = null;
			}
			try
			{
				XlChartType chartType2 = series.ChartType;
				if ((uint)(chartType2 - 57) <= 1u)
				{
					TargetBand.B(chart, range, range2, d, e);
				}
				else
				{
					A(chart, range, range2, d, e);
				}
				try
				{
					chart.ChartArea.Select();
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				Forms.ErrorMessage(ex8.Message);
				clsReporting.LogException(ex8);
				ProjectData.ClearProjectError();
			}
			if (num != -4142.0)
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
				Axis axis2 = (Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue);
				if (axis2.MaximumScale != num)
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
					axis2.MaximumScale = num;
				}
				axis2 = null;
			}
			application.AutoCorrect.AutoExpandListRange = autoExpandListRange;
			application.ScreenUpdating = true;
			application = null;
			chart = null;
			series = null;
			range = null;
			Range range3 = null;
			range2 = null;
			Range range4 = null;
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(68933));
			return;
		}
		IL_010b:
		text = VH.A(68660);
		goto IL_011a;
	}

	private static void A(Chart A, Range B, Range C, bool D, QD E)
	{
		if (D)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					TargetBand.A(A, XlChartType.xlColumnStacked, B, C, E);
					return;
				}
			}
		}
		TargetBand.B(A, XlChartType.xlColumnStacked, B, C, E);
	}

	private static void B(Chart A, Range B, Range C, bool D, QD E)
	{
		if (D)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					TargetBand.A(A, XlChartType.xlBarStacked, B, C, E);
					return;
				}
			}
		}
		TargetBand.B(A, XlChartType.xlBarStacked, B, C, E);
	}

	private static void A(Chart A, XlChartType B, Range C, Range D, QD E)
	{
		Application application = A.Application;
		Range range = null;
		List<XlChartType> b = TargetBand.A(A);
		Range range2 = ((Range)C.Rows[C.Rows.Count, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)3, RuntimeHelpers.GetObjectValue(Missing.Value));
		range2.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
		range2 = range2.get_Offset((object)(-3), (object)0);
		if (D != null)
		{
			range = ((Range)D.Rows[D.Rows.Count, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)3, RuntimeHelpers.GetObjectValue(Missing.Value));
			range.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
			range = range.get_Offset((object)(-3), (object)0);
		}
		NewLateBinding.LateSetComplex(range2.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(57302), new object[1] { E.A }, null, null, OptimisticSet: false, RValueBase: true);
		XlChartType chartType = A.ChartType;
		if (chartType <= XlChartType.xlBarStacked)
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
			if (chartType != XlChartType.xlColumnStacked)
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
				if (chartType != XlChartType.xlBarStacked)
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
					goto IL_04cd;
				}
			}
		}
		else if (chartType != XlChartType.xlLineStacked && chartType != XlChartType.xlLineMarkersStacked)
		{
			goto IL_04cd;
		}
		NewLateBinding.LateSetComplex(range2.Cells[2, 1], null, VH.A(68956), new object[1] { Operators.ConcatenateObject(VH.A(48936) + Conversions.ToString(E.B) + VH.A(13778), NewLateBinding.LateGet(range2.Cells[1, 1], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)) }, null, null, OptimisticSet: false, RValueBase: true);
		List<string> list = new List<string>();
		Range range3 = JH.A(C, (Application)null);
		if (range3 != null)
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
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = range3.Columns.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range4 = (Range)enumerator.Current;
					list.Add(VH.A(54414) + range4.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(39904));
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
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			range3 = null;
		}
		NewLateBinding.LateSetComplex(range2.Cells[3, 1], null, VH.A(68956), new object[1] { Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(68971) + string.Join(VH.A(2378), list.ToArray()) + VH.A(68994), NewLateBinding.LateGet(range2.Cells[2, 1], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)), VH.A(13778)), NewLateBinding.LateGet(range2.Cells[1, 1], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)), VH.A(39904)) }, null, null, OptimisticSet: false, RValueBase: true);
		list = null;
		((Range)range2.Rows[2, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Resize((object)2, RuntimeHelpers.GetObjectValue(Missing.Value)).FillRight();
		goto IL_0746;
		IL_0746:
		string right = TargetBand.A();
		int num = 1;
		do
		{
			string text = TargetBand.A(num);
			Series series;
			if (range != null)
			{
				NewLateBinding.LateSetComplex(range.Rows[num, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(57302), new object[1] { text }, null, null, OptimisticSet: false, RValueBase: true);
				object instance = A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value));
				string memberName = VH.A(60813);
				object[] array = new object[4];
				object[] array2 = array;
				Application application2 = application;
				object instance2 = range.Rows[num, RuntimeHelpers.GetObjectValue(Missing.Value)];
				string memberName2 = VH.A(5814);
				object[] array3 = new object[3];
				array3[1] = 0;
				array3[2] = 0;
				array3[0] = true;
				object left = Operators.ConcatenateObject(NewLateBinding.LateGet(instance2, null, memberName2, array3, new string[1] { VH.A(68999) }, null, null), right);
				object instance3 = range2.Rows[num, RuntimeHelpers.GetObjectValue(Missing.Value)];
				string memberName3 = VH.A(5814);
				object[] array4 = new object[3];
				array4[1] = 0;
				array4[2] = 0;
				array4[0] = true;
				array2[0] = ((_Application)application2).get_Range(Operators.ConcatenateObject(left, NewLateBinding.LateGet(instance3, null, memberName3, array4, new string[1] { VH.A(68999) }, null, null)), RuntimeHelpers.GetObjectValue(Missing.Value));
				array[1] = XlRowCol.xlRows;
				array[2] = true;
				array[3] = false;
				series = (Series)NewLateBinding.LateGet(instance, null, memberName, array, new string[4]
				{
					VH.A(69016),
					VH.A(69029),
					VH.A(69042),
					VH.A(69067)
				}, null, null);
			}
			else
			{
				object instance4 = A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value));
				string memberName4 = VH.A(60813);
				object[] array = new object[4];
				object[] array5 = array;
				Application application3 = application;
				object instance5 = range2.Rows[num, RuntimeHelpers.GetObjectValue(Missing.Value)];
				string memberName5 = VH.A(5814);
				object[] array6 = new object[3];
				array6[1] = 0;
				array6[2] = 0;
				array6[0] = true;
				array5[0] = ((_Application)application3).get_Range(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance5, null, memberName5, array6, new string[1] { VH.A(68999) }, null, null)), RuntimeHelpers.GetObjectValue(Missing.Value));
				array[1] = XlRowCol.xlRows;
				array[2] = false;
				array[3] = false;
				series = (Series)NewLateBinding.LateGet(instance4, null, memberName4, array, new string[4]
				{
					VH.A(69016),
					VH.A(69029),
					VH.A(69042),
					VH.A(69067)
				}, null, null);
				series.Name = text;
			}
			series.AxisGroup = XlAxisGroup.xlPrimary;
			series.ChartType = B;
			if (num != 1)
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
				if (num == 2)
				{
					TargetBand.A(series, E);
				}
				else
				{
					series.Format.Fill.Visible = MsoTriState.msoFalse;
				}
			}
			else
			{
				if (E.A < 0f)
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
					if (E.B > 0f)
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
						TargetBand.A(series, E);
						goto IL_0aee;
					}
				}
				series.Format.Fill.Visible = MsoTriState.msoFalse;
			}
			goto IL_0aee;
			IL_0aee:
			num = checked(num + 1);
		}
		while (num <= 3);
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			TargetBand.A(A, b, B);
			TargetBand.A(range2);
			application = null;
			Series series = null;
			range2 = null;
			range = null;
			b = null;
			return;
		}
		IL_04cd:
		Range range5 = JH.A(range2, (Application)null);
		if (range5 != null)
		{
			IEnumerator enumerator2 = default(IEnumerator);
			try
			{
				enumerator2 = range5.Columns.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					Range range6 = (Range)enumerator2.Current;
					NewLateBinding.LateSetComplex(range6.Rows[2, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(68956), new object[1] { Operators.ConcatenateObject(VH.A(48936) + Conversions.ToString(E.B) + VH.A(13778), NewLateBinding.LateGet(range6.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)) }, null, null, OptimisticSet: false, RValueBase: true);
					NewLateBinding.LateSetComplex(range6.Rows[3, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(68956), new object[1] { Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(68971) + C.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(68994), NewLateBinding.LateGet(range6.Rows[2, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)), VH.A(13778)), NewLateBinding.LateGet(range6.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)), VH.A(39904)) }, null, null, OptimisticSet: false, RValueBase: true);
				}
			}
			finally
			{
				if (enumerator2 is IDisposable)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						(enumerator2 as IDisposable).Dispose();
						break;
					}
				}
			}
			range5 = null;
		}
		goto IL_0746;
	}

	private static void B(Chart A, XlChartType B, Range C, Range D, QD E)
	{
		Application application = A.Application;
		Range range = null;
		List<XlChartType> b = TargetBand.A(A);
		Range range2 = ((Range)C.Columns[C.Columns.Count, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)0, (object)1).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)3);
		range2.Insert(XlInsertShiftDirection.xlShiftToRight, RuntimeHelpers.GetObjectValue(Missing.Value));
		range2 = range2.get_Offset((object)0, (object)(-3));
		if (D != null)
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
			range = ((Range)D.Columns[D.Columns.Count, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)0, (object)1).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)3);
			range.Insert(XlInsertShiftDirection.xlShiftToRight, RuntimeHelpers.GetObjectValue(Missing.Value));
			range = range.get_Offset((object)0, (object)(-3));
		}
		NewLateBinding.LateSetComplex(range2.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(57302), new object[1] { E.A }, null, null, OptimisticSet: false, RValueBase: true);
		XlChartType chartType = A.ChartType;
		if (chartType <= XlChartType.xlBarStacked)
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
			if (chartType != XlChartType.xlColumnStacked)
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
				if (chartType != XlChartType.xlBarStacked)
				{
					goto IL_04e1;
				}
			}
		}
		else if (chartType != XlChartType.xlLineStacked)
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
			if (chartType != XlChartType.xlLineMarkersStacked)
			{
				goto IL_04e1;
			}
		}
		NewLateBinding.LateSetComplex(range2.Cells[1, 2], null, VH.A(68956), new object[1] { Operators.ConcatenateObject(VH.A(48936) + Conversions.ToString(E.B) + VH.A(13778), NewLateBinding.LateGet(range2.Cells[1, 1], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)) }, null, null, OptimisticSet: false, RValueBase: true);
		List<string> list = new List<string>();
		Range range3 = JH.A(C, (Application)null);
		if (range3 != null)
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = range3.Rows.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range4 = (Range)enumerator.Current;
					list.Add(VH.A(54414) + range4.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(39904));
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_0334;
					}
					continue;
					end_IL_0334:
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
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			range3 = null;
		}
		NewLateBinding.LateSetComplex(range2.Cells[1, 3], null, VH.A(68956), new object[1] { Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(68971) + string.Join(VH.A(2378), list.ToArray()) + VH.A(68994), NewLateBinding.LateGet(range2.Cells[1, 2], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)), VH.A(13778)), NewLateBinding.LateGet(range2.Cells[1, 1], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)), VH.A(39904)) }, null, null, OptimisticSet: false, RValueBase: true);
		list = null;
		((Range)range2.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)2).FillDown();
		goto IL_077c;
		IL_077c:
		string right = TargetBand.A();
		int num = 1;
		Series series;
		do
		{
			string text = TargetBand.A(num);
			if (range != null)
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
				NewLateBinding.LateSetComplex(range.Columns[num, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(57302), new object[1] { text }, null, null, OptimisticSet: false, RValueBase: true);
				object instance = A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value));
				string memberName = VH.A(60813);
				object[] array = new object[4];
				object[] array2 = array;
				Application application2 = application;
				object instance2 = range.Columns[num, RuntimeHelpers.GetObjectValue(Missing.Value)];
				string memberName2 = VH.A(5814);
				object[] array3 = new object[3];
				array3[1] = 0;
				array3[2] = 0;
				array3[0] = true;
				object left = Operators.ConcatenateObject(NewLateBinding.LateGet(instance2, null, memberName2, array3, new string[1] { VH.A(68999) }, null, null), right);
				object instance3 = range2.Columns[num, RuntimeHelpers.GetObjectValue(Missing.Value)];
				string memberName3 = VH.A(5814);
				object[] array4 = new object[3];
				array4[1] = 0;
				array4[2] = 0;
				array4[0] = true;
				array2[0] = ((_Application)application2).get_Range(Operators.ConcatenateObject(left, NewLateBinding.LateGet(instance3, null, memberName3, array4, new string[1] { VH.A(68999) }, null, null)), RuntimeHelpers.GetObjectValue(Missing.Value));
				array[1] = XlRowCol.xlColumns;
				array[2] = true;
				array[3] = false;
				series = (Series)NewLateBinding.LateGet(instance, null, memberName, array, new string[4]
				{
					VH.A(69016),
					VH.A(69029),
					VH.A(69042),
					VH.A(69067)
				}, null, null);
			}
			else
			{
				object instance4 = A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value));
				string memberName4 = VH.A(60813);
				object[] array = new object[4];
				object[] array5 = array;
				Application application3 = application;
				object instance5 = range2.Columns[num, RuntimeHelpers.GetObjectValue(Missing.Value)];
				string memberName5 = VH.A(5814);
				object[] array6 = new object[3];
				array6[1] = 0;
				array6[2] = 0;
				array6[0] = true;
				array5[0] = ((_Application)application3).get_Range(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance5, null, memberName5, array6, new string[1] { VH.A(68999) }, null, null)), RuntimeHelpers.GetObjectValue(Missing.Value));
				array[1] = XlRowCol.xlColumns;
				array[2] = false;
				array[3] = false;
				series = (Series)NewLateBinding.LateGet(instance4, null, memberName4, array, new string[4]
				{
					VH.A(69016),
					VH.A(69029),
					VH.A(69042),
					VH.A(69067)
				}, null, null);
				series.Name = text;
			}
			series.AxisGroup = XlAxisGroup.xlPrimary;
			series.ChartType = B;
			if (num != 1)
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
				if (num != 2)
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
					series.Format.Fill.Visible = MsoTriState.msoFalse;
				}
				else
				{
					TargetBand.A(series, E);
				}
			}
			else
			{
				if (E.A < 0f)
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
					if (E.B > 0f)
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
						TargetBand.A(series, E);
						goto IL_0b36;
					}
				}
				series.Format.Fill.Visible = MsoTriState.msoFalse;
			}
			goto IL_0b36;
			IL_0b36:
			num = checked(num + 1);
		}
		while (num <= 3);
		TargetBand.A(A, b, B);
		TargetBand.A(range2);
		application = null;
		series = null;
		range2 = null;
		range = null;
		b = null;
		return;
		IL_04e1:
		Range range5 = JH.A(range2, (Application)null);
		if (range5 != null)
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
			IEnumerator enumerator2 = default(IEnumerator);
			try
			{
				enumerator2 = range5.Rows.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					Range range6 = (Range)enumerator2.Current;
					NewLateBinding.LateSetComplex(range6.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(68956), new object[1] { Operators.ConcatenateObject(VH.A(48936) + Conversions.ToString(E.B) + VH.A(13778), NewLateBinding.LateGet(range6.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)) }, null, null, OptimisticSet: false, RValueBase: true);
					NewLateBinding.LateSetComplex(range6.Columns[3, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(68956), new object[1] { Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(68971) + C.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(68994), NewLateBinding.LateGet(range6.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)), VH.A(13778)), NewLateBinding.LateGet(range6.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)), VH.A(39904)) }, null, null, OptimisticSet: false, RValueBase: true);
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_074d;
					}
					continue;
					end_IL_074d:
					break;
				}
			}
			finally
			{
				if (enumerator2 is IDisposable)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						(enumerator2 as IDisposable).Dispose();
						break;
					}
				}
			}
			range5 = null;
		}
		goto IL_077c;
	}

	private static string A(int A)
	{
		if (A != 1)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (A != 2)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								return VH.A(69118);
							}
						}
					}
					return VH.A(69107);
				}
			}
		}
		return VH.A(69096);
	}

	private static List<XlChartType> A(Chart A)
	{
		List<XlChartType> list = new List<XlChartType>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = ((IEnumerable)A.FullSeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
			while (enumerator.MoveNext())
			{
				Series series = (Series)enumerator.Current;
				list.Add(series.ChartType);
			}
			while (true)
			{
				switch (6)
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
		return list;
	}

	private static void A(Chart A, List<XlChartType> B, XlChartType C)
	{
		checked
		{
			int num = B.Count - 1;
			for (int i = 0; i <= num; i++)
			{
				Series series = (Series)A.FullSeriesCollection(i + 1);
				if (!series.IsFiltered)
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
					series.AxisGroup = XlAxisGroup.xlSecondary;
					series.ChartType = B[i];
				}
				series = null;
			}
			((Series)A.SeriesCollection(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(52690), new object[0], null, null, null)))).ChartType = C;
			((ChartGroup)A.ChartGroups(1)).GapWidth = 0;
			((_Chart)A).set_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlValue, (object)XlAxisGroup.xlSecondary, (object)false);
		}
	}

	private static void A(Range A)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Cells.GetEnumerator();
			while (enumerator.MoveNext())
			{
				AutoColor.AutoColorIfNotEmpty((Range)enumerator.Current);
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				return;
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
	}

	private static string A()
	{
		return CultureInfo.CurrentCulture.TextInfo.ListSeparator;
	}

	private static void A(Series A, QD B)
	{
		Microsoft.Office.Interop.Excel.FillFormat fill = A.Format.Fill;
		fill.ForeColor.RGB = B.A;
		fill.Transparency = (float)((double)B.B / 100.0);
		_ = null;
	}

	private static QD A(Chart A, ref bool B)
	{
		QD qD = default(QD);
		try
		{
			XmlDocument A2 = KH.A.SettingsXml;
			qD = TargetBand.A(ref A2);
			try
			{
				Axis obj = (Axis)A.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue);
				double maximumScale = obj.MaximumScale;
				double minimumScale = obj.MinimumScale;
				_ = null;
				if (!((double)qD.B > maximumScale))
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
					if (qD.B != 0f)
					{
						goto IL_007d;
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
				qD.B = (float)maximumScale;
				goto IL_007d;
				IL_007d:
				if ((double)qD.A < minimumScale)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						qD.A = (float)minimumScale;
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
			wpfTargetBand wpfTargetBand2 = new wpfTargetBand();
			System.Drawing.Color color = ColorTranslator.FromOle(qD.A);
			wpfTargetBand2.btnColor.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(color.R, color.G, color.B));
			wpfTargetBand2.numTransparency.Value = qD.B;
			wpfTargetBand2.txtLower.Text = qD.A.ToString();
			wpfTargetBand2.txtUpper.Text = qD.B.ToString();
			wpfTargetBand2.ShowDialog();
			if (wpfTargetBand2.DialogResult.HasValue && wpfTargetBand2.DialogResult.Value)
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
				System.Windows.Media.Color color2 = ((SolidColorBrush)wpfTargetBand2.btnColor.Foreground).Color;
				qD = new QD
				{
					A = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(color2.R, color2.G, color2.B)),
					B = checked((int)Math.Round(wpfTargetBand2.numTransparency.Value.Value)),
					A = Conversions.ToSingle(wpfTargetBand2.txtLower.Text),
					B = Conversions.ToSingle(wpfTargetBand2.txtUpper.Text)
				};
				TargetBand.A(qD, ref A2);
			}
			else
			{
				B = true;
			}
			wpfTargetBand2 = null;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.WarningMessage(VH.A(69125));
			qD.A = 49407;
			qD.B = 50;
			qD.A = 0f;
			qD.B = 100f;
			ProjectData.ClearProjectError();
		}
		finally
		{
			XmlDocument A2 = null;
		}
		return qD;
	}

	private static QD A(ref XmlDocument A)
	{
		QD result = default(QD);
		XmlDocument xmlDocument = A;
		result.A = clsColors.RGB2Ole(xmlDocument.SelectSingleNode(VH.A(69267)).InnerText);
		result.B = checked((int)Math.Round(float.Parse(xmlDocument.SelectSingleNode(VH.A(69304)).InnerText, CultureInfo.InvariantCulture)));
		result.A = float.Parse(xmlDocument.SelectSingleNode(VH.A(69355)).InnerText, CultureInfo.InvariantCulture);
		result.B = float.Parse(xmlDocument.SelectSingleNode(VH.A(69398)).InnerText, CultureInfo.InvariantCulture);
		xmlDocument = null;
		return result;
	}

	private static void A(QD A, ref XmlDocument B)
	{
		B.SelectSingleNode(VH.A(69267)).InnerText = clsColors.Color2RGB(ColorTranslator.FromOle(A.A));
		B.SelectSingleNode(VH.A(69304)).InnerText = A.B.ToString(CultureInfo.InvariantCulture);
		B.SelectSingleNode(VH.A(69355)).InnerText = A.A.ToString(CultureInfo.InvariantCulture);
		B.SelectSingleNode(VH.A(69398)).InnerText = A.B.ToString(CultureInfo.InvariantCulture);
		KH.A.SaveSettings(B);
	}
}
