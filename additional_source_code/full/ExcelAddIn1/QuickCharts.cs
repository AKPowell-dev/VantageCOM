using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

public sealed class QuickCharts
{
	private enum XF
	{
		A = 0,
		B = 1,
		C = 2,
		D = 3,
		E = 4,
		F = 5,
		G = 6,
		H = 7,
		I = 8,
		J = 10,
		K = 11,
		L = 12
	}

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

	public static readonly int OPTIONS_DARK_YELLOW = ColorTranslator.ToOle(Color.FromArgb(191, 144, 0));

	public static readonly int OPTIONS_TABLE_FILL = ColorTranslator.ToOle(Color.FromArgb(255, 242, 204));

	public static readonly int GRIDLINES_COLOR = ColorTranslator.ToOle(Color.FromArgb(217, 217, 217));

	private static readonly int m_A = 4;

	private static List<StandardSize> m_A;

	private static List<StandardSize> A
	{
		get
		{
			if (QuickCharts.m_A == null)
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
				QuickCharts.m_A = clsPublish.GetStandardSizes();
			}
			return QuickCharts.m_A;
		}
	}

	public static ChartObject AddChart(Worksheet ws, float sngWidth, float sngHeight)
	{
		ChartSize chartSize = default(ChartSize);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
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
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			seriesCollection = null;
			return;
		}
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
			result = ColorTranslator.ToOle(Color.Blue);
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
				switch (2)
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
				switch (3)
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
				switch (2)
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
				switch (4)
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
		if (dblMax == 0.0)
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
			if (dblMin == 0.0)
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
				dblMax = 1.0;
			}
		}
		double num2 = Math.Log(dblMax - dblMin) / Math.Log(10.0);
		double num3 = Math.Pow(10.0, num2 - (double)checked((int)Math.Round(num2)));
		double num4 = num3;
		double num5;
		if (num4 >= 0.0 && num4 <= 2.5)
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
			num3 = 0.2;
			num5 = 0.05;
		}
		else
		{
			if (num4 >= 2.5)
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
				if (num4 <= 5.0)
				{
					num3 = 0.5;
					num5 = 0.1;
					goto IL_0275;
				}
			}
			if (num4 >= 5.0 && num4 <= 7.5)
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
			}
			else
			{
				num3 = 2.0;
				num5 = 0.5;
			}
		}
		goto IL_0275;
		IL_0275:
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
		Validation validation = rng.Validation;
		validation.InputMessage = VH.A(71154);
		validation.ErrorMessage = validation.InputMessage;
		validation.ErrorTitle = VH.A(40448);
		_ = null;
	}

	public static void AddNumFormatValidation(Range rng)
	{
		rng.Validation.Add(XlDVType.xlValidateInputOnly, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		Validation validation = rng.Validation;
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
		rng.Font.Color = ColorTranslator.ToOle(Color.Blue);
		_ = null;
	}

	public static void SetChartWidth(ref XmlDocument xmlDoc, decimal decWidth)
	{
		xmlDoc.SelectSingleNode(VH.A(141262)).InnerText = decWidth.ToString();
	}

	public static void SetChartHeight(ref XmlDocument xmlDoc, decimal decHeight)
	{
		xmlDoc.SelectSingleNode(VH.A(141311)).InnerText = decHeight.ToString();
	}

	public static bool GetPreserveFormulas(XmlDocument xmlDoc)
	{
		return Conversions.ToBoolean(xmlDoc.SelectSingleNode(VH.A(141362)).InnerText);
	}

	public static void SetPreserveFormulas(ref XmlDocument xmlDoc, bool blnPreserve)
	{
		xmlDoc.SelectSingleNode(VH.A(141362)).InnerText = (0 - (blnPreserve ? 1 : 0)).ToString();
	}

	public static int GetGapWidth(XmlDocument xmlDoc)
	{
		return Conversions.ToInteger(xmlDoc.SelectSingleNode(VH.A(141423)).InnerText);
	}

	public static void SetGapWidth(ref XmlDocument xmlDoc, int intGap)
	{
		xmlDoc.SelectSingleNode(VH.A(141423)).InnerText = intGap.ToString();
	}

	public static string GetLineColor(XmlDocument xmlDoc)
	{
		return xmlDoc.SelectSingleNode(VH.A(141468)).InnerText;
	}

	public static void SetLineColor(ref XmlDocument xmlDoc, string strRGB)
	{
		xmlDoc.SelectSingleNode(VH.A(141468)).InnerText = strRGB;
	}

	public static string LoadCommonSettings(XmlDocument xmlDoc, NumericUpDown numWidth, NumericUpDown numHeight, Label lblWidth, Label lblHeight)
	{
		string text = xmlDoc.SelectSingleNode(VH.A(141262)).InnerText;
		string text2 = xmlDoc.SelectSingleNode(VH.A(141311)).InnerText;
		string text3 = clsPublish.SystemDecimalSeparator();
		if (!text2.Contains(text3))
		{
			text2 = clsPublish.ConvertToSystemDecimal(text2, text3);
			text = clsPublish.ConvertToSystemDecimal(text, text3);
		}
		string text4;
		if (!RegionInfo.CurrentRegion.IsMetric)
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
			text4 = VH.A(71870);
		}
		else
		{
			text4 = VH.A(71875);
		}
		string result = (lblHeight.Text = (lblWidth.Text = text4));
		numWidth.Value = Conversions.ToDecimal(text);
		numHeight.Value = Conversions.ToDecimal(text2);
		return result;
	}

	public static void ChartDialogLoad(ComboBox cbx, NumericUpDown numHeight, NumericUpDown numWidth, MeasurementUnits units)
	{
		//IL_0031: Unknown result type (might be due to invalid IL or missing references)
		//IL_0036: Unknown result type (might be due to invalid IL or missing references)
		//IL_0043: Unknown result type (might be due to invalid IL or missing references)
		//IL_0048: Unknown result type (might be due to invalid IL or missing references)
		List<StandardSize> standardSizes = clsPublish.GetStandardSizes();
		ComboBox comboBox = cbx;
		comboBox.SelectedIndex = 0;
		checked
		{
			int num = comboBox.Items.Count - 1;
			int num2 = 1;
			while (true)
			{
				if (num2 <= num)
				{
					float num3 = standardSizes[num2 - 1].Height;
					float num4 = standardSizes[num2 - 1].Width;
					if (units == MeasurementUnits.Centimeters)
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
						num3 = (float)Math.Round(num3 * clsPublish.CENTIMETERS_PER_INCH, QuickCharts.m_A);
						num4 = (float)Math.Round(num4 * clsPublish.CENTIMETERS_PER_INCH, QuickCharts.m_A);
					}
					if (Convert.ToSingle(numHeight.Value) == num3)
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
						if (Convert.ToSingle(numWidth.Value) == num4)
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
							comboBox.SelectedIndex = num2;
							break;
						}
					}
					num2++;
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
				break;
			}
			comboBox = null;
		}
	}

	public static void StandardSizeSelectedIndexChanged(ComboBox cbx, NumericUpDown numHeight, NumericUpDown numWidth, MeasurementUnits units)
	{
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_0068: Unknown result type (might be due to invalid IL or missing references)
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		int selectedIndex = cbx.SelectedIndex;
		if (selectedIndex > 0)
		{
			StandardSize val = QuickCharts.A[checked(selectedIndex - 1)];
			if (units == MeasurementUnits.Inches)
			{
				numHeight.Value = new decimal(val.Height);
				numWidth.Value = new decimal(val.Width);
			}
			else
			{
				numHeight.Value = new decimal(Math.Round(val.Height * clsPublish.CENTIMETERS_PER_INCH, QuickCharts.m_A));
				numWidth.Value = new decimal(Math.Round(val.Width * clsPublish.CENTIMETERS_PER_INCH, QuickCharts.m_A));
			}
		}
	}

	public static void CheckForStandardSize(ComboBox cbx, NumericUpDown numHeight, NumericUpDown numWidth)
	{
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0060: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Unknown result type (might be due to invalid IL or missing references)
		float num = Convert.ToSingle(numHeight.Value);
		float num2 = Convert.ToSingle(numWidth.Value);
		checked
		{
			int num3 = QuickCharts.A.Count - 1;
			cbx.BeginUpdate();
			cbx.SelectedIndex = 0;
			int num4 = num3;
			for (int i = 0; i <= num4; i++)
			{
				if (num != QuickCharts.A[i].Height || num2 != QuickCharts.A[i].Width)
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
				cbx.SelectedIndex = i + 1;
				break;
			}
			cbx.EndUpdate();
		}
	}

	private static void A(object A, CancelEventArgs B)
	{
		if (int.TryParse(((TextBox)A).Text, out var result))
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
			if (result <= 500 && result >= 0)
			{
				return;
			}
		}
		Forms.WarningMessage(VH.A(71880));
		B.Cancel = true;
	}

	private static void A(object A, EventArgs B)
	{
		Button button = (Button)A;
		ColorDialog colorDialog = new ColorDialog();
		colorDialog.AnyColor = true;
		colorDialog.SolidColorOnly = true;
		colorDialog.AllowFullOpen = true;
		colorDialog.FullOpen = true;
		colorDialog.ShowHelp = false;
		colorDialog.Color = button.ForeColor;
		if (colorDialog.ShowDialog() == DialogResult.OK)
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
			button.ForeColor = colorDialog.Color;
			button.Image = clsColors.ColorSquare(colorDialog.Color, button);
		}
		colorDialog.Dispose();
		colorDialog = null;
		button = null;
	}

	public static void HandleGapWidthChange(TextBox txt)
	{
		txt.Validating += A;
	}

	public static void HandleColorChange(Button btn)
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

	public static Range SampleChartData(TextBox txt, Func<Worksheet, Range> f)
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
}
