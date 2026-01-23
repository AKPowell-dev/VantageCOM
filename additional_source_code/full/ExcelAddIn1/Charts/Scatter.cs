using System;
using System.Collections;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class Scatter
{
	private struct FE
	{
		public float A;

		public float B;

		public XlDataLabelPosition A;

		public bool A;

		public bool B;

		public bool C;

		public int A;

		public bool D;
	}

	private static readonly string m_A = VH.A(78949);

	public static void Create()
	{
		if (!Licensing.AllowQuickChartOperation())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		bool B = true;
		checked
		{
			if (application.Selection is Range)
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
				if (!Workbooks.IsShared(application.ActiveWorkbook, true, (System.Windows.Window)null))
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
					Range A = (Range)application.Selection;
					FE fE = Scatter.A(ref A, ref B);
					if (B)
					{
						if (Operators.ConditionalCompareObjectNotEqual(A.Columns.CountLarge, 3, TextCompare: false))
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
							Forms.WarningMessage(VH.A(78515));
							B = false;
						}
						else
						{
							IEnumerator enumerator = default(IEnumerator);
							try
							{
								enumerator = ((_Worksheet)A.Worksheet).get_Range(RuntimeHelpers.GetObjectValue(A.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)]), RuntimeHelpers.GetObjectValue(A.Columns[3, RuntimeHelpers.GetObjectValue(Missing.Value)])).Cells.GetEnumerator();
								while (true)
								{
									if (enumerator.MoveNext())
									{
										Range range = (Range)enumerator.Current;
										if (!Operators.ConditionalCompareObjectEqual(range.Formula, string.Empty, TextCompare: false))
										{
											while (true)
											{
												switch (7)
												{
												case 0:
													break;
												default:
													goto end_IL_0153;
												}
												continue;
												end_IL_0153:
												break;
											}
											if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range.Value2)))
											{
												continue;
											}
										}
										Forms.WarningMessage(VH.A(78636));
										B = false;
										break;
									}
									while (true)
									{
										switch (6)
										{
										case 0:
											break;
										default:
											goto end_IL_0197;
										}
										continue;
										end_IL_0197:
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
						}
						if (B)
						{
							XlCalculation calc = default(XlCalculation);
							QuickCharts2.PrepareExcel(application, ref calc);
							ChartObject chartObject;
							try
							{
								QuickCharts2.ChartSize chartSize = default(QuickCharts2.ChartSize);
								if (!RegionInfo.CurrentRegion.IsMetric)
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
									chartSize.Width = clsPublish.InchesToPoints(fE.A);
									chartSize.Height = clsPublish.InchesToPoints(fE.B);
								}
								else
								{
									chartSize.Width = clsPublish.CentimetersToPoints(fE.A);
									chartSize.Height = clsPublish.CentimetersToPoints(fE.B);
								}
								ChartObjects chartObjects = (ChartObjects)A.Worksheet.ChartObjects(RuntimeHelpers.GetObjectValue(Missing.Value));
								int count = chartObjects.Count;
								try
								{
									chartObject = chartObjects.Add(Conversions.ToDouble(A.Left), Conversions.ToDouble(Operators.AddObject(Operators.AddObject(A.Top, A.Height), 10)), chartSize.Width, chartSize.Height);
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									ChartObjects chartObjects2 = chartObjects;
									if (chartObjects2.Count == count + 1)
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
										NewLateBinding.LateCall(chartObjects2.Item(chartObjects2.Count), null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
									}
									chartObjects2 = null;
									chartObject = (ChartObject)A.Worksheet.Shapes.AddChart2(RuntimeHelpers.GetObjectValue(Missing.Value), XlChartType.xlXYScatter, RuntimeHelpers.GetObjectValue(A.Left), Operators.AddObject(Operators.AddObject(A.Top, A.Height), 10), chartSize.Width, chartSize.Height, RuntimeHelpers.GetObjectValue(Missing.Value)).Chart.Parent;
									ProjectData.ClearProjectError();
								}
								chartObjects = null;
								QuickCharts2.DeleteAllSeries(chartObject);
								chartObject.Placement = XlPlacement.xlFreeFloating;
								_ = chartObject.BottomRightCell.Row + 2;
								Chart chart = chartObject.Chart;
								chart.ChartType = XlChartType.xlXYScatter;
								QuickCharts2.RequireAxes(chart);
								Series series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add((Range)A.Columns[3, RuntimeHelpers.GetObjectValue(Missing.Value)], XlRowCol.xlColumns, false, false, RuntimeHelpers.GetObjectValue(Missing.Value));
								series.XValues = VH.A(48936) + ((Range)A.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
								if (KH.A.ChartSeriesColors.Any())
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
									int markerBackgroundColor = (series.MarkerForegroundColor = clsColors.RGB2Ole(KH.A.ChartSeriesColors[0]));
									series.MarkerBackgroundColor = markerBackgroundColor;
								}
								series.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								((Microsoft.Office.Interop.Excel.DataLabels)series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Position = fE.A;
								if (Conversion.Val(application.Version) < 15.0)
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
									application.ScreenUpdating = true;
									application.ScreenUpdating = false;
								}
								DataLabels.LinkToRange((Microsoft.Office.Interop.Excel.DataLabels)series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), (Range)A.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)]);
								if (fE.A)
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
									Trendline trendline = ((Trendlines)series.Trendlines(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(XlTrendlineType.xlLinear, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									if (fE.B)
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
										trendline.DisplayRSquared = true;
										if (fE.C)
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
											trendline.DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = fE.A;
										}
									}
									trendline.Format.Line.ForeColor.RGB = fE.A;
									trendline = null;
								}
								series = null;
								QuickCharts2.AxisScale axisScale = QuickCharts2.GetAxisScale(application.WorksheetFunction.Min(RuntimeHelpers.GetObjectValue(A.Columns[3, RuntimeHelpers.GetObjectValue(Missing.Value)]), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), application.WorksheetFunction.Max(RuntimeHelpers.GetObjectValue(A.Columns[3, RuntimeHelpers.GetObjectValue(Missing.Value)]), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
								Axis axis = (Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue);
								if (axis.MinimumScale > axisScale.MaximumScale)
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
									axis.MaximumScale = axisScale.MaximumScale;
									axis.MinimumScale = axisScale.MinimumScale;
								}
								else
								{
									axis.MinimumScale = axisScale.MinimumScale;
									axis.MaximumScale = axisScale.MaximumScale;
								}
								axis.MajorUnit = axisScale.MajorUnit;
								axis = null;
								axisScale = QuickCharts2.GetAxisScale(application.WorksheetFunction.Min(RuntimeHelpers.GetObjectValue(A.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)]), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), application.WorksheetFunction.Max(RuntimeHelpers.GetObjectValue(A.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)]), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
								Axis axis2 = (Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory);
								if (axis2.MinimumScale > axisScale.MaximumScale)
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
									axis2.MaximumScale = axisScale.MaximumScale;
									axis2.MinimumScale = axisScale.MinimumScale;
								}
								else
								{
									axis2.MinimumScale = axisScale.MinimumScale;
									axis2.MaximumScale = axisScale.MaximumScale;
								}
								axis2.MajorUnit = axisScale.MajorUnit;
								axis2 = null;
								QuickCharts2.CleanUpChart(chart);
								Chart chart2 = chart;
								((Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory)).HasMajorGridlines = true;
								((Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory)).MajorGridlines.Format.Line.ForeColor.RGB = QuickCharts2.GRIDLINES_COLOR;
								chart2.ChartArea.Format.Line.Visible = MsoTriState.msoFalse;
								chart2.HasLegend = false;
								chart2.ChartArea.Select();
								chart2 = null;
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								Forms.ErrorMessage(ex4.Message);
								clsReporting.LogException(ex4);
								ProjectData.ClearProjectError();
							}
							QuickCharts2.RestoreExcel(application, calc);
							chartObject = null;
							QuickCharts2.LogActivity(VH.A(78733));
						}
					}
					A = null;
				}
			}
			application = null;
		}
	}

	private static FE A(ref Range A, ref bool B)
	{
		XmlDocument xmlDoc = KH.A.SettingsXml;
		wpfScatter wpfScatter2 = new wpfScatter();
		QuickCharts2.HandleColorChange(wpfScatter2.btnTrendlineColor);
		wpfScatter2.Range = A;
		wpfScatter2.txtAddress.Text = A.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		QuickCharts2.LoadCommonSettings(xmlDoc, wpfScatter2.numChartWidth, wpfScatter2.numChartHeight);
		wpfScatter2.chkColorLabel.IsChecked = Conversions.ToBoolean(xmlDoc.SelectSingleNode(Scatter.m_A + VH.A(78760)).InnerText);
		wpfScatter2.chkShowR2.IsChecked = Conversions.ToBoolean(xmlDoc.SelectSingleNode(Scatter.m_A + VH.A(78785)).InnerText);
		wpfScatter2.chkShowTrendline.IsChecked = Conversions.ToBoolean(xmlDoc.SelectSingleNode(Scatter.m_A + VH.A(78798)).InnerText);
		bool? isChecked = wpfScatter2.chkShowTrendline.IsChecked;
		bool? flag;
		if (!isChecked.HasValue)
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
			flag = isChecked;
		}
		else
		{
			flag = isChecked != true;
		}
		isChecked = flag;
		if (isChecked == true)
		{
			wpfScatter2.chkShowR2.IsChecked = false;
			wpfScatter2.chkShowR2.IsEnabled = false;
			wpfScatter2.chkColorLabel.IsChecked = false;
			wpfScatter2.chkColorLabel.IsEnabled = false;
		}
		else
		{
			isChecked = wpfScatter2.chkShowR2.IsChecked;
			bool? flag2;
			if (!isChecked.HasValue)
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
				flag2 = isChecked;
			}
			else
			{
				flag2 = isChecked != true;
			}
			isChecked = flag2;
			if (isChecked == true)
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
				wpfScatter2.chkColorLabel.IsChecked = false;
				wpfScatter2.chkColorLabel.IsEnabled = false;
			}
		}
		wpfScatter2.chkScaleAxes.IsChecked = Conversions.ToBoolean(xmlDoc.SelectSingleNode(Scatter.m_A + VH.A(78825)).InnerText);
		QuickCharts2.SetButtonColor(wpfScatter2.btnTrendlineColor, QuickCharts2.GetLineColor(xmlDoc));
		switch ((XlDataLabelPosition)Conversions.ToInteger(xmlDoc.SelectSingleNode(Scatter.m_A + VH.A(55388)).InnerText))
		{
		case XlDataLabelPosition.xlLabelPositionAbove:
			wpfScatter2.cbxLabelPosn.SelectedIndex = 0;
			break;
		case XlDataLabelPosition.xlLabelPositionBelow:
			wpfScatter2.cbxLabelPosn.SelectedIndex = 1;
			break;
		case XlDataLabelPosition.xlLabelPositionRight:
			wpfScatter2.cbxLabelPosn.SelectedIndex = 2;
			break;
		default:
			wpfScatter2.cbxLabelPosn.SelectedIndex = 3;
			break;
		}
		wpfScatter2.ShowDialog();
		FE result = default(FE);
		if (wpfScatter2.DialogResult.HasValue)
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
			if (wpfScatter2.DialogResult.Value)
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
				A = wpfScatter2.Range;
				result = new FE
				{
					A = (float)wpfScatter2.numChartWidth.Value.Value,
					B = (float)wpfScatter2.numChartHeight.Value.Value,
					A = wpfScatter2.chkShowTrendline.IsChecked.Value,
					B = wpfScatter2.chkShowR2.IsChecked.Value,
					C = wpfScatter2.chkColorLabel.IsChecked.Value,
					D = wpfScatter2.chkScaleAxes.IsChecked.Value
				};
				System.Windows.Media.Color buttonColor = QuickCharts2.GetButtonColor(wpfScatter2.btnTrendlineColor);
				result.A = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(buttonColor.R, buttonColor.G, buttonColor.B));
				switch (wpfScatter2.cbxLabelPosn.SelectedIndex)
				{
				case 0:
					result.A = XlDataLabelPosition.xlLabelPositionAbove;
					break;
				case 1:
					result.A = XlDataLabelPosition.xlLabelPositionBelow;
					break;
				case 2:
					result.A = XlDataLabelPosition.xlLabelPositionRight;
					break;
				case 3:
					result.A = XlDataLabelPosition.xlLabelPositionLeft;
					break;
				}
				QuickCharts2.SetChartWidth(ref xmlDoc, new decimal(wpfScatter2.numChartWidth.Value.Value));
				QuickCharts2.SetChartHeight(ref xmlDoc, new decimal(wpfScatter2.numChartHeight.Value.Value));
				QuickCharts2.SetLineColor(ref xmlDoc, QuickCharts2.Color2RGB(buttonColor));
				xmlDoc.SelectSingleNode(Scatter.m_A + VH.A(78798)).InnerText = (0 - (wpfScatter2.chkShowTrendline.IsChecked.Value ? 1 : 0)).ToString();
				xmlDoc.SelectSingleNode(Scatter.m_A + VH.A(78785)).InnerText = (0 - (wpfScatter2.chkShowR2.IsChecked.Value ? 1 : 0)).ToString();
				xmlDoc.SelectSingleNode(Scatter.m_A + VH.A(78760)).InnerText = (0 - (wpfScatter2.chkColorLabel.IsChecked.Value ? 1 : 0)).ToString();
				xmlDoc.SelectSingleNode(Scatter.m_A + VH.A(78825)).InnerText = (0 - (wpfScatter2.chkScaleAxes.IsChecked.Value ? 1 : 0)).ToString();
				XmlNode xmlNode = xmlDoc.SelectSingleNode(Scatter.m_A + VH.A(55388));
				int a = (int)result.A;
				xmlNode.InnerText = a.ToString();
				KH.A.SaveSettings(xmlDoc);
				goto IL_065e;
			}
		}
		B = false;
		goto IL_065e;
		IL_065e:
		wpfScatter2 = null;
		xmlDoc = null;
		return result;
	}

	public static Range Example(Worksheet ws)
	{
		Worksheet worksheet = ws;
		((_Worksheet)worksheet).get_Range((object)VH.A(78844), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78849);
		((_Worksheet)worksheet).get_Range((object)VH.A(76929), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78862);
		((_Worksheet)worksheet).get_Range((object)VH.A(76965), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78869);
		((_Worksheet)worksheet).get_Range((object)VH.A(76981), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78876);
		((_Worksheet)worksheet).get_Range((object)VH.A(76997), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78883);
		((_Worksheet)worksheet).get_Range((object)VH.A(78364), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78890);
		((_Worksheet)worksheet).get_Range((object)VH.A(78390), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78897);
		((_Worksheet)worksheet).get_Range((object)VH.A(76877), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78904);
		((_Worksheet)worksheet).get_Range((object)VH.A(76945), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 6.03;
		((_Worksheet)worksheet).get_Range((object)VH.A(61417), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 8.34;
		((_Worksheet)worksheet).get_Range((object)VH.A(61422), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 7.01;
		((_Worksheet)worksheet).get_Range((object)VH.A(61427), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 6.3;
		((_Worksheet)worksheet).get_Range((object)VH.A(61439), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 8.75;
		((_Worksheet)worksheet).get_Range((object)VH.A(61451), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 10.13;
		((_Worksheet)worksheet).get_Range((object)VH.A(57617), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78921);
		((_Worksheet)worksheet).get_Range((object)VH.A(76950), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 10.28;
		((_Worksheet)worksheet).get_Range((object)VH.A(61486), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 12.75;
		((_Worksheet)worksheet).get_Range((object)VH.A(61502), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 9.79;
		((_Worksheet)worksheet).get_Range((object)VH.A(61507), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 10.62;
		((_Worksheet)worksheet).get_Range((object)VH.A(61512), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 12.02;
		((_Worksheet)worksheet).get_Range((object)VH.A(61517), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 12.45;
		Range range = ((_Worksheet)worksheet).get_Range((object)VH.A(78444), RuntimeHelpers.GetObjectValue(Missing.Value));
		try
		{
			range.NumberFormat = QuickCharts2.CURRENCY_FORMAT_2;
			range.Font.Color = QuickCharts2.InputColor();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		range = null;
		Range result = ((_Worksheet)worksheet).get_Range((object)VH.A(78938), RuntimeHelpers.GetObjectValue(Missing.Value));
		worksheet = null;
		return result;
	}
}
