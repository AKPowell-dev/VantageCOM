using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Xml;
using A;
using Foo.Controls;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class QuickCharts2
{
	public struct ChartSize
	{
		public double Width;

		public double Height;
	}

	public struct AxisScale
	{
		public double MinimumScale;

		public double MaximumScale;

		public double MinorUnit;

		public double MajorUnit;
	}

	public enum MeasurementUnits
	{
		Inches,
		Centimeters
	}

	public static readonly string CURRENCY_FORMAT_1 = VH.A(72054);

	public static readonly string CURRENCY_FORMAT_2 = VH.A(72117);

	public static readonly string NUMFORMAT_HIDDEN = VH.A(2545);

	public static readonly int OPTIONS_DARK_YELLOW = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 144, 0));

	public static readonly int OPTIONS_TABLE_FILL = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204));

	public static readonly int GRIDLINES_COLOR = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(217, 217, 217));

	private static readonly int m_A = 4;

	private static List<StandardSize> m_A = null;

	private static List<StandardSize> A
	{
		get
		{
			if (QuickCharts2.m_A == null)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				QuickCharts2.m_A = clsPublish.GetStandardSizes();
			}
			return QuickCharts2.m_A;
		}
	}

	public static ChartObject AddChart(Worksheet ws, float sngWidth, float sngHeight)
	{
		ChartSize chartSize = default(ChartSize);
		if (!RegionInfo.CurrentRegion.IsMetric)
		{
			chartSize.Width = clsPublish.InchesToPoints(sngWidth);
			chartSize.Height = clsPublish.InchesToPoints(sngHeight);
		}
		else
		{
			chartSize.Width = clsPublish.CentimetersToPoints(sngWidth);
			chartSize.Height = clsPublish.CentimetersToPoints(sngHeight);
		}
		ChartObject chartObject = Charts.AddChart(ws, chartSize.Width, chartSize.Height);
		DeleteAllSeries(chartObject);
		return chartObject;
	}

	public static void DeleteAllSeries(ChartObject chtObj)
	{
		SeriesCollection seriesCollection = (SeriesCollection)chtObj.Chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value));
		for (int i = seriesCollection.Count; i >= 1; i = checked(i + -1))
		{
			seriesCollection.Item(i).Delete();
		}
		seriesCollection = null;
	}

	public static void RequireAxes(Chart cht)
	{
		((_Chart)cht).set_HasAxis((object)XlAxisType.xlValue, (object)XlAxisGroup.xlPrimary, (object)true);
		((_Chart)cht).set_HasAxis((object)XlAxisType.xlCategory, (object)XlAxisGroup.xlPrimary, (object)true);
	}

	public static void PrepareExcel(Microsoft.Office.Interop.Excel.Application xlApp, ref XlCalculation calc)
	{
		calc = xlApp.Calculation;
		xlApp.Calculation = XlCalculation.xlCalculationManual;
		xlApp.ScreenUpdating = false;
		xlApp.EnableEvents = false;
	}

	public static void RestoreExcel(Microsoft.Office.Interop.Excel.Application xlApp, XlCalculation calc)
	{
		xlApp.Calculation = calc;
		xlApp.EnableEvents = true;
		xlApp.ScreenUpdating = true;
	}

	public static void LogActivity(string strActivity)
	{
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, strActivity);
	}

	public static int DefaultColor()
	{
		int result;
		try
		{
			result = KH.A.DefaultFontColor;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = 0;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static int InputColor()
	{
		int result;
		try
		{
			result = clsColors.RGB2Ole(KH.A.AutoColors[0]);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = ColorTranslator.ToOle(System.Drawing.Color.Blue);
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static int LinkColor()
	{
		int result;
		try
		{
			result = clsColors.RGB2Ole(KH.A.AutoColors[3]);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = clsColors.RGB2Ole(VH.A(71139));
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static AxisScale GetAxisScale(double dblMin, double dblMax)
	{
		AxisScale result = default(AxisScale);
		if (dblMax == dblMin)
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
			dblMax *= 1.01;
			dblMin *= 0.99;
		}
		if (dblMax < dblMin)
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
			double num = dblMax;
			dblMax = dblMin;
			dblMin = num;
		}
		if (dblMax > 0.0)
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
			dblMax += (dblMax - dblMin) * 0.01;
		}
		else if (dblMax < 0.0)
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
			dblMax = Math.Min(dblMax + (dblMax - dblMin) * 0.01, 0.0);
		}
		else
		{
			dblMax = 0.0;
		}
		if (dblMin > 0.0)
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
			dblMin = Math.Max(dblMin - (dblMax - dblMin) * 0.01, 0.0);
		}
		else if (dblMin < 0.0)
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
			dblMin -= (dblMax - dblMin) * 0.01;
		}
		else
		{
			dblMin = 0.0;
		}
		if (dblMax == 0.0 && dblMin == 0.0)
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
			dblMax = 1.0;
		}
		double num2 = Math.Log(dblMax - dblMin) / Math.Log(10.0);
		double num3 = Math.Pow(10.0, num2 - (double)checked((int)Math.Round(num2)));
		double num4 = num3;
		double num5;
		if (num4 >= 0.0)
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
			if (num4 <= 2.5)
			{
				num3 = 0.2;
				num5 = 0.05;
				goto IL_027f;
			}
		}
		if (num4 >= 2.5)
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
			if (num4 <= 5.0)
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
				num3 = 0.5;
				num5 = 0.1;
				goto IL_027f;
			}
		}
		if (num4 >= 5.0)
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
			if (num4 <= 7.5)
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
				num3 = 1.0;
				num5 = 0.2;
				goto IL_027f;
			}
		}
		num3 = 2.0;
		num5 = 0.5;
		goto IL_027f;
		IL_027f:
		checked
		{
			num3 *= Math.Pow(10.0, (int)Math.Round(num2));
			num5 *= Math.Pow(10.0, (int)Math.Round(num2));
			result.MinimumScale = num3 * (double)((int)Math.Round(dblMin / num3) - 1);
			result.MaximumScale = num3 * (double)((int)Math.Round(dblMax / num3) + 1);
			result.MajorUnit = num3;
			result.MinorUnit = num5;
			return result;
		}
	}

	public static void CleanUpChart(Chart cht)
	{
		try
		{
			((Axis)cht.Axes(XlAxisType.xlValue)).MajorGridlines.Format.Line.ForeColor.RGB = GRIDLINES_COLOR;
			((Axis)cht.Axes(XlAxisType.xlValue)).MajorTickMark = XlTickMark.xlTickMarkNone;
			((Axis)cht.Axes(XlAxisType.xlCategory)).MajorTickMark = XlTickMark.xlTickMarkNone;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void AddPercentageValidation(Range rng)
	{
		rng.Validation.Add(XlDVType.xlValidateDecimal, XlDVAlertStyle.xlValidAlertInformation, XlFormatConditionOperator.xlBetween, 0, 0.25);
		Microsoft.Office.Interop.Excel.Validation validation = rng.Validation;
		validation.InputMessage = VH.A(71154);
		validation.ErrorMessage = validation.InputMessage;
		validation.ErrorTitle = VH.A(40448);
		_ = null;
	}

	public static void AddNumFormatValidation(Range rng)
	{
		rng.Validation.Add(XlDVType.xlValidateInputOnly, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		Microsoft.Office.Interop.Excel.Validation validation = rng.Validation;
		validation.InputMessage = VH.A(71300);
		validation.ShowError = false;
		_ = null;
	}

	public static void FormatOptionsHeader(Range rng, string strText)
	{
		rng.Font.Color = OPTIONS_DARK_YELLOW;
		rng.Font.Bold = true;
		rng.Value2 = strText;
		rng.EntireColumn.ColumnWidth = 18;
		_ = null;
	}

	public static void FormatOptionsInput(Range rng)
	{
		rng.Interior.Color = clsColors.RGB2Ole(VH.A(71600));
		rng.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Blue);
		_ = null;
	}

	public static void SetChartWidth(ref XmlDocument xmlDoc, decimal decWidth)
	{
		xmlDoc.DocumentElement.SelectSingleNode(VH.A(71623)).InnerText = decWidth.ToString();
	}

	public static void SetChartHeight(ref XmlDocument xmlDoc, decimal decHeight)
	{
		xmlDoc.DocumentElement.SelectSingleNode(VH.A(71668)).InnerText = decHeight.ToString();
	}

	public static bool GetPreserveFormulas(XmlDocument xmlDoc)
	{
		return Conversions.ToBoolean(xmlDoc.DocumentElement.SelectSingleNode(VH.A(71715)).InnerText);
	}

	public static void SetPreserveFormulas(ref XmlDocument xmlDoc, bool blnPreserve)
	{
		xmlDoc.DocumentElement.SelectSingleNode(VH.A(71715)).InnerText = (0 - (blnPreserve ? 1 : 0)).ToString();
	}

	public static int GetGapWidth(XmlDocument xmlDoc)
	{
		return Conversions.ToInteger(xmlDoc.DocumentElement.SelectSingleNode(VH.A(71772)).InnerText);
	}

	public static void SetGapWidth(ref XmlDocument xmlDoc, int intGap)
	{
		xmlDoc.DocumentElement.SelectSingleNode(VH.A(71772)).InnerText = intGap.ToString();
	}

	public static string GetLineColor(XmlDocument xmlDoc)
	{
		return xmlDoc.DocumentElement.SelectSingleNode(VH.A(71813)).InnerText;
	}

	public static void SetLineColor(ref XmlDocument xmlDoc, string strRGB)
	{
		xmlDoc.DocumentElement.SelectSingleNode(VH.A(71813)).InnerText = strRGB;
	}

	public static void LoadCommonSettings(XmlDocument xmlDoc, MacNumericUpDown numWidth, MacNumericUpDown numHeight)
	{
		string text = xmlDoc.DocumentElement.SelectSingleNode(VH.A(71623)).InnerText;
		string text2 = xmlDoc.DocumentElement.SelectSingleNode(VH.A(71668)).InnerText;
		string text3 = clsPublish.SystemDecimalSeparator();
		if (!text2.Contains(text3))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			text2 = clsPublish.ConvertToSystemDecimal(text2, text3);
			text = clsPublish.ConvertToSystemDecimal(text, text3);
		}
		string text4;
		if (!RegionInfo.CurrentRegion.IsMetric)
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
			text4 = VH.A(71870);
		}
		else
		{
			text4 = VH.A(71875);
		}
		string customUnit = (numHeight.CustomUnit = text4);
		numWidth.CustomUnit = customUnit;
		numWidth.Value = Convert.ToDouble(Conversions.ToDecimal(text));
		numHeight.Value = Convert.ToDouble(Conversions.ToDecimal(text2));
	}

	public static void ChartDialogLoad(System.Windows.Controls.ComboBox cbx, MacNumericUpDown numHeight, MacNumericUpDown numWidth)
	{
		//IL_00ae: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c9: Unknown result type (might be due to invalid IL or missing references)
		List<StandardSizeItem> list = new List<StandardSizeItem>();
		_ = QuickCharts2.A;
		list.Add(new StandardSizeItem());
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = clsPublish.StandardSizeNodes().GetEnumerator();
			while (enumerator.MoveNext())
			{
				XmlNode nd = (XmlNode)enumerator.Current;
				list.Add(new StandardSizeItem(nd));
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
					switch (5)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		cbx.ItemsSource = list;
		cbx.SelectedIndex = 0;
		checked
		{
			int num = QuickCharts2.A.Count - 1;
			int num2 = 0;
			while (true)
			{
				if (num2 <= num)
				{
					float num3 = QuickCharts2.A[num2].Height;
					float num4 = QuickCharts2.A[num2].Width;
					if (RegionInfo.CurrentRegion.IsMetric)
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
						num3 = (float)Math.Round(num3 * clsPublish.CENTIMETERS_PER_INCH, QuickCharts2.m_A);
						num4 = (float)Math.Round(num4 * clsPublish.CENTIMETERS_PER_INCH, QuickCharts2.m_A);
					}
					if ((float)numHeight.Value.Value == num3)
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
						if ((float)numWidth.Value.Value == num4)
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
							cbx.SelectedIndex = num2 + 1;
							break;
						}
					}
					num2++;
					continue;
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
				break;
			}
			list = null;
		}
	}

	public static void StandardSizeSelectedIndexChanged(System.Windows.Controls.ComboBox cbx, MacNumericUpDown numHeight, MacNumericUpDown numWidth)
	{
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		//IL_001b: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0067: Unknown result type (might be due to invalid IL or missing references)
		//IL_008b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0042: Unknown result type (might be due to invalid IL or missing references)
		//IL_0054: Unknown result type (might be due to invalid IL or missing references)
		int selectedIndex = cbx.SelectedIndex;
		if (selectedIndex <= 0)
		{
			return;
		}
		StandardSize val = QuickCharts2.A[checked(selectedIndex - 1)];
		if (!RegionInfo.CurrentRegion.IsMetric)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					numHeight.Value = val.Height;
					numWidth.Value = val.Width;
					return;
				}
			}
		}
		numHeight.Value = Math.Round(val.Height * clsPublish.CENTIMETERS_PER_INCH, QuickCharts2.m_A);
		numWidth.Value = Math.Round(val.Width * clsPublish.CENTIMETERS_PER_INCH, QuickCharts2.m_A);
	}

	public static void CheckForStandardSize(System.Windows.Controls.ComboBox cbx, MacNumericUpDown numHeight, MacNumericUpDown numWidth)
	{
		//IL_004b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0050: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		//IL_0079: Unknown result type (might be due to invalid IL or missing references)
		float num = (float)numHeight.Value.Value;
		float num2 = (float)numWidth.Value.Value;
		checked
		{
			int num3 = QuickCharts2.A.Count - 1;
			cbx.SelectedIndex = 0;
			int num4 = num3;
			for (int i = 0; i <= num4; i++)
			{
				if (num != QuickCharts2.A[i].Height)
				{
					continue;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (num2 != QuickCharts2.A[i].Width)
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
					cbx.SelectedIndex = i + 1;
					return;
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
	}

	private static void A(object A, RoutedEventArgs B)
	{
		System.Windows.Controls.TextBox textBox = (System.Windows.Controls.TextBox)A;
		if (int.TryParse(textBox.Text, out var result))
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
			if (result <= 500)
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
				if (result >= 0)
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
			}
		}
		Forms.WarningMessage(VH.A(71880));
		textBox.Focus();
		textBox.SelectAll();
	}

	private static void A(object A, EventArgs B)
	{
		System.Windows.Controls.Button button = (System.Windows.Controls.Button)A;
		System.Windows.Shapes.Rectangle rectangle = (System.Windows.Shapes.Rectangle)button.Content;
		if (rectangle == null)
		{
			ColorDialog colorDialog = new ColorDialog();
			colorDialog.AnyColor = true;
			colorDialog.SolidColorOnly = true;
			colorDialog.AllowFullOpen = true;
			colorDialog.FullOpen = true;
			colorDialog.ShowHelp = false;
			System.Windows.Media.Color color = ((SolidColorBrush)button.Foreground).Color;
			colorDialog.Color = System.Drawing.Color.FromArgb(color.R, color.G, color.B);
			if (colorDialog.ShowDialog() == DialogResult.OK)
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
				color = System.Windows.Media.Color.FromRgb(colorDialog.Color.R, colorDialog.Color.G, colorDialog.Color.B);
				button.Foreground = new SolidColorBrush(color);
			}
			colorDialog.Dispose();
			colorDialog = null;
		}
		else
		{
			ColorDialog colorDialog2 = new ColorDialog();
			colorDialog2.AnyColor = true;
			colorDialog2.SolidColorOnly = true;
			colorDialog2.AllowFullOpen = true;
			colorDialog2.FullOpen = true;
			colorDialog2.ShowHelp = false;
			System.Windows.Media.Color color = ((SolidColorBrush)rectangle.Fill).Color;
			colorDialog2.Color = System.Drawing.Color.FromArgb(color.R, color.G, color.B);
			if (colorDialog2.ShowDialog() == DialogResult.OK)
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
				color = System.Windows.Media.Color.FromRgb(colorDialog2.Color.R, colorDialog2.Color.G, colorDialog2.Color.B);
				rectangle.Fill = new SolidColorBrush(color);
			}
			colorDialog2.Dispose();
			colorDialog2 = null;
		}
		button = null;
		rectangle = null;
	}

	public static void HandleGapWidthChange(System.Windows.Controls.TextBox txt)
	{
		txt.LostFocus += A;
	}

	public static void HandleColorChange(System.Windows.Controls.Button btn)
	{
		btn.Click += A;
	}

	public static void DeleteSampleWorksheet(Worksheet ws)
	{
		if (ws == null)
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
			Microsoft.Office.Interop.Excel.Application application = ws.Application;
			application.DisplayAlerts = false;
			try
			{
				ws.Delete();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			application.DisplayAlerts = true;
			application = null;
			return;
		}
	}

	public static Range SampleChartData(System.Windows.Controls.TextBox txt, Func<Worksheet, Range> f)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application.ActiveWorkbook;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		Range range = default(Range);
		Worksheet arg;
		try
		{
			arg = (Worksheet)activeWorkbook.Worksheets.Add(RuntimeHelpers.GetObjectValue(activeWorkbook.ActiveSheet), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			range = f(arg);
			txt.Text = range.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			range.Select();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
		application.ScreenUpdating = true;
		application.EnableEvents = true;
		application = null;
		activeWorkbook = null;
		arg = null;
		return range;
	}

	public static void SetButtonColor(System.Windows.Controls.Button btn, string strRGB)
	{
		btn.Foreground = new SolidColorBrush(A(strRGB));
	}

	public static System.Windows.Media.Color GetButtonColor(System.Windows.Controls.Button btn)
	{
		return ((SolidColorBrush)btn.Foreground).Color;
	}

	private static System.Windows.Media.Color A(string A)
	{
		System.Windows.Media.Color result;
		try
		{
			string[] array = Strings.Split(A, VH.A(2378));
			result = checked(System.Windows.Media.Color.FromRgb((byte)Conversions.ToInteger(array[0]), (byte)Conversions.ToInteger(array[1]), (byte)Conversions.ToInteger(array[2])));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = Colors.Transparent;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static string Color2RGB(System.Windows.Media.Color clr)
	{
		if (clr == Colors.Transparent)
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
					return "";
				}
			}
		}
		return Conversions.ToString((int)clr.R) + VH.A(2378) + Conversions.ToString((int)clr.G) + VH.A(2378) + Conversions.ToString((int)clr.B);
	}
}
