using System;
using System.Collections;
using System.Drawing;
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

public sealed class FootballField
{
	private enum LE
	{
		A,
		B
	}

	private struct ME
	{
		public float A;

		public float B;

		public bool A;

		public int A;

		public LE A;

		public int B;

		public MsoLineDashStyle A;
	}

	private struct NE
	{
		public Range A;

		public Range B;

		public Range C;

		public Range D;

		public Range E;

		public Range F;

		public Range G;

		public Range H;
	}

	private static readonly string m_A = VH.A(82247);

	private static readonly int m_A = 1;

	private static readonly int m_B = 2;

	private static readonly int C = 3;

	private static readonly int D = 4;

	private static readonly int E = 15;

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
					ME mE = FootballField.A(ref A, ref B);
					if (B)
					{
						if (Operators.ConditionalCompareObjectNotEqual(A.Columns.CountLarge, 3, TextCompare: false))
						{
							Forms.WarningMessage(VH.A(81237));
							B = false;
						}
						else
						{
							IEnumerator enumerator = default(IEnumerator);
							try
							{
								enumerator = ((_Worksheet)A.Worksheet).get_Range(RuntimeHelpers.GetObjectValue(A.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)]), RuntimeHelpers.GetObjectValue(A.Columns[3, RuntimeHelpers.GetObjectValue(Missing.Value)])).Cells.GetEnumerator();
								while (enumerator.MoveNext())
								{
									Range range = (Range)enumerator.Current;
									if (!Operators.ConditionalCompareObjectEqual(range.Formula, string.Empty, TextCompare: false))
									{
										while (true)
										{
											switch (6)
											{
											case 0:
												break;
											default:
												goto end_IL_0147;
											}
											continue;
											end_IL_0147:
											break;
										}
										if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range.Value2)))
										{
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
									}
									Forms.WarningMessage(VH.A(78636));
									B = false;
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
						}
						if (B)
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
							int num = QuickCharts2.DefaultColor();
							int e = QuickCharts2.InputColor();
							int f = QuickCharts2.LinkColor();
							XlCalculation calc = default(XlCalculation);
							QuickCharts2.PrepareExcel(application, ref calc);
							ChartObject chartObject;
							Chart chart;
							Range range2;
							try
							{
								Worksheet worksheet = (Worksheet)application.ActiveWorkbook.Worksheets.Add(A.Worksheet, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								chartObject = QuickCharts2.AddChart(worksheet, mE.A, mE.B);
								chartObject.Placement = XlPlacement.xlFreeFloating;
								int val = chartObject.BottomRightCell.Row + 2;
								chart = chartObject.Chart;
								QuickCharts2.RequireAxes(chart);
								NE nE = FootballField.A(worksheet, chartObject);
								val = Math.Max(val, E + 3);
								FootballField.A(mE, worksheet, ref A, val, e, f);
								int num2 = Conversions.ToInteger(Operators.AddObject(A.Rows.CountLarge, 1));
								range2 = ((Range)A.Cells[1, 2]).get_Offset((object)(-1), (object)0).get_Resize((object)num2, (object)5);
								A.Select();
								Range range3 = range2;
								string formula = VH.A(57636) + nE.D.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(79182) + ((Range)range3.Columns[5, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + nE.E.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75669);
								string formula2 = VH.A(57636) + nE.D.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(79182) + ((Range)range3.Columns[6, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + nE.F.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75669);
								string formula3 = VH.A(57636) + nE.D.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(79182) + ((Range)range3.Columns[7, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + nE.G.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75669);
								string text = ((Range)range3.Columns[FootballField.m_A, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + ((Range)range3.Columns[FootballField.m_B, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								((Range)range3.Cells[1, FootballField.m_A]).Value2 = VH.A(81373);
								((Range)range3.Cells[1, FootballField.m_B]).Value2 = VH.A(81380);
								((Range)range3.Cells[1, C]).Value2 = VH.A(76918);
								((Range)range3.Cells[1, D]).Value2 = VH.A(57387);
								((Range)range3.Cells[1, 5]).Value2 = VH.A(72816);
								if (mE.A == LE.A)
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
									((Range)range3.Columns[5, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)).Formula = VH.A(57636) + nE.A.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(81389) + text + VH.A(81408);
									((Range)range3.Cells[1, 6]).Value2 = VH.A(72846);
									((Range)range3.Columns[6, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)).Formula = VH.A(57636) + nE.B.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(81423) + text + VH.A(81408);
									((Range)range3.Cells[1, 7]).Value2 = VH.A(81440);
									((Range)range3.Columns[7, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)).Formula = VH.A(57636) + nE.C.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + nE.H.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75241);
									((Range)range3.Cells[1, 8]).Value2 = VH.A(81451);
									((Range)range3.Cells[1, 9]).Value2 = VH.A(81478);
									((Range)range3.Cells[1, 10]).Value2 = VH.A(81503);
									((Range)range3.Cells[RuntimeHelpers.GetObjectValue(range3.Rows.CountLarge), 8]).Formula = formula;
									((Range)range3.Cells[RuntimeHelpers.GetObjectValue(range3.Rows.CountLarge), 9]).Formula = formula2;
									((Range)range3.Cells[RuntimeHelpers.GetObjectValue(range3.Rows.CountLarge), 10]).Formula = formula3;
								}
								else
								{
									((Range)range3.Columns[5, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)).Formula = VH.A(57636) + nE.A.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(81389) + text + VH.A(81408);
									((Range)range3.Cells[1, 6]).Value2 = VH.A(72846);
									((Range)range3.Columns[6, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)).Formula = VH.A(57636) + nE.B.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(81423) + text + VH.A(81408);
									((Range)range3.Cells[1, 7]).Value2 = VH.A(81440);
									((Range)range3.Columns[7, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)).Formula = VH.A(57636) + nE.C.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + nE.H.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75241);
									((Range)range3.Cells[1, 8]).Value2 = VH.A(81451);
									((Range)range3.Cells[1, 9]).Value2 = VH.A(81478);
									((Range)range3.Cells[1, 10]).Value2 = VH.A(81503);
									((Range)range3.Cells[3, 8]).Formula = formula;
									((Range)range3.Cells[3, 9]).Formula = formula2;
									((Range)range3.Cells[3, 10]).Formula = formula3;
								}
								int num3 = num2;
								for (int i = 2; i <= num3; i++)
								{
									((Range)range3.Cells[i, C]).Formula = VH.A(48936) + ((Range)range3.Cells[i, FootballField.m_B]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(13778) + ((Range)range3.Cells[i, FootballField.m_A]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									((Range)range3.Cells[i, D]).Formula = VH.A(72965) + ((Range)range3.Cells[2, 1]).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(72945) + ((Range)range3.Cells[i - 1, D]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(72958);
								}
								int num4 = num2;
								for (int j = 2; j <= num4; j++)
								{
									string numberFormat = Conversions.ToString(((Range)range3.Cells[j, FootballField.m_A]).NumberFormat);
									Range obj = (Range)range3.Cells[j, C];
									obj.NumberFormat = numberFormat;
									obj.Font.Color = num;
									_ = null;
									Range obj2 = (Range)range3.Cells[j, D];
									obj2.NumberFormat = VH.A(81526);
									obj2.Font.Color = num;
									_ = null;
									Range obj3 = (Range)range3.Cells[j, 5];
									obj3.NumberFormat = numberFormat;
									obj3.Font.Color = num;
									_ = null;
									Range obj4 = (Range)range3.Cells[j, 6];
									obj4.NumberFormat = numberFormat;
									obj4.Font.Color = num;
									_ = null;
									Range obj5 = (Range)range3.Cells[j, 7];
									obj5.NumberFormat = numberFormat;
									obj5.Font.Color = num;
									_ = null;
								}
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									range3 = null;
									if (mE.A == LE.B)
									{
										FootballField.A(chart, range2, A, mE, num2);
									}
									else
									{
										FootballField.B(chart, range2, A, mE, num2);
									}
									QuickCharts2.AxisScale axisScale = QuickCharts2.GetAxisScale(application.WorksheetFunction.Min(((Range)range2.Columns[FootballField.m_A, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), application.WorksheetFunction.Max(((Range)range2.Columns[FootballField.m_B, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
									Axis axis = (Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue);
									if (axis.MinimumScale > axisScale.MaximumScale)
									{
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
									QuickCharts2.CleanUpChart(chart);
									Chart chart2 = chart;
									if (mE.A == LE.A)
									{
										((Axis)chart2.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory)).TickLabelPosition = XlTickLabelPosition.xlTickLabelPositionLow;
									}
									((ChartGroup)chart2.ChartGroups(1)).GapWidth = mE.A;
									chart2.HasLegend = false;
									chart2.ChartArea.Format.Line.Visible = MsoTriState.msoFalse;
									chart2.ChartArea.Select();
									chart2 = null;
									break;
								}
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								Forms.ErrorMessage(ex2.Message);
								clsReporting.LogException(ex2);
								ProjectData.ClearProjectError();
							}
							QuickCharts2.RestoreExcel(application, calc);
							chartObject = null;
							chart = null;
							range2 = null;
							QuickCharts2.LogActivity(VH.A(81533));
						}
					}
					A = null;
				}
			}
			application = null;
		}
	}

	private static ME A(ref Range A, ref bool B)
	{
		XmlDocument xmlDoc = KH.A.SettingsXml;
		wpfFootballField wpfFootballField2 = new wpfFootballField();
		QuickCharts2.HandleColorChange(wpfFootballField2.btnColorLine);
		wpfFootballField2.Range = A;
		wpfFootballField2.txtAddress.Text = A.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		QuickCharts2.LoadCommonSettings(xmlDoc, wpfFootballField2.numChartWidth, wpfFootballField2.numChartHeight);
		wpfFootballField2.numGapWidth.Value = QuickCharts2.GetGapWidth(xmlDoc);
		wpfFootballField2.chkPreserveFormulas.IsChecked = QuickCharts2.GetPreserveFormulas(xmlDoc);
		LE lE = (LE)Conversions.ToInteger(xmlDoc.SelectSingleNode(FootballField.m_A + VH.A(60421)).InnerText);
		if (lE == LE.B)
		{
			wpfFootballField2.radBars.IsChecked = true;
		}
		else
		{
			wpfFootballField2.radColumns.IsChecked = true;
		}
		MsoLineDashStyle msoLineDashStyle = (MsoLineDashStyle)Conversions.ToInteger(xmlDoc.SelectSingleNode(FootballField.m_A + VH.A(81574)).InnerText);
		int selectedIndex = default(int);
		if (msoLineDashStyle != MsoLineDashStyle.msoLineSolid)
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
			switch (msoLineDashStyle)
			{
			case MsoLineDashStyle.msoLineSysDot:
				selectedIndex = 1;
				break;
			case MsoLineDashStyle.msoLineSysDash:
				selectedIndex = 2;
				break;
			case MsoLineDashStyle.msoLineDash:
				selectedIndex = 3;
				break;
			case MsoLineDashStyle.msoLineLongDash:
				selectedIndex = 4;
				break;
			}
		}
		else
		{
			selectedIndex = 0;
		}
		wpfFootballField2.cbxLineStyle.SelectedIndex = selectedIndex;
		QuickCharts2.SetButtonColor(wpfFootballField2.btnColorLine, QuickCharts2.GetLineColor(xmlDoc));
		wpfFootballField2.ShowDialog();
		ME result = default(ME);
		if (wpfFootballField2.DialogResult.HasValue)
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
			if (wpfFootballField2.DialogResult.Value)
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
				A = wpfFootballField2.Range;
				XmlNode xmlNode;
				checked
				{
					result = new ME
					{
						A = wpfFootballField2.chkPreserveFormulas.IsChecked.Value,
						A = (float)wpfFootballField2.numChartWidth.Value.Value,
						B = (float)wpfFootballField2.numChartHeight.Value.Value,
						A = (int)Math.Round(wpfFootballField2.numGapWidth.Value.Value)
					};
					if (wpfFootballField2.radBars.IsChecked == true)
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
						result.A = LE.B;
					}
					else
					{
						result.A = LE.A;
					}
					switch (wpfFootballField2.cbxLineStyle.SelectedIndex)
					{
					case 0:
						result.A = MsoLineDashStyle.msoLineSolid;
						break;
					case 1:
						result.A = MsoLineDashStyle.msoLineSysDot;
						break;
					case 2:
						result.A = MsoLineDashStyle.msoLineSysDash;
						break;
					case 3:
						result.A = MsoLineDashStyle.msoLineDash;
						break;
					case 4:
						result.A = MsoLineDashStyle.msoLineLongDash;
						break;
					}
					System.Windows.Media.Color buttonColor = QuickCharts2.GetButtonColor(wpfFootballField2.btnColorLine);
					result.B = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(buttonColor.R, buttonColor.G, buttonColor.B));
					QuickCharts2.SetChartWidth(ref xmlDoc, new decimal(wpfFootballField2.numChartWidth.Value.Value));
					QuickCharts2.SetChartHeight(ref xmlDoc, new decimal(wpfFootballField2.numChartHeight.Value.Value));
					QuickCharts2.SetGapWidth(ref xmlDoc, (int)Math.Round(wpfFootballField2.numGapWidth.Value.Value));
					QuickCharts2.SetPreserveFormulas(ref xmlDoc, wpfFootballField2.chkPreserveFormulas.IsChecked.Value);
					QuickCharts2.SetLineColor(ref xmlDoc, QuickCharts2.Color2RGB(buttonColor));
					xmlNode = xmlDoc.SelectSingleNode(FootballField.m_A + VH.A(60421));
				}
				int a = (int)result.A;
				xmlNode.InnerText = a.ToString();
				XmlNode xmlNode2 = xmlDoc.SelectSingleNode(FootballField.m_A + VH.A(81574));
				a = (int)result.A;
				xmlNode2.InnerText = a.ToString();
				KH.A.SaveSettings(xmlDoc);
				goto IL_0485;
			}
		}
		B = false;
		goto IL_0485;
		IL_0485:
		wpfFootballField2 = null;
		xmlDoc = null;
		return result;
	}

	private static void A(ME A, Worksheet B, ref Range C, int D, int E, int F)
	{
		Microsoft.Office.Interop.Excel.Application application = C.Application;
		int num = Conversions.ToInteger(C.Rows.CountLarge);
		int num2 = Conversions.ToInteger(C.Columns.CountLarge);
		checked
		{
			if (A.A)
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
				C.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
				((Range)B.Cells[D, 1]).PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				int num3 = num;
				for (int i = 1; i <= num3; i++)
				{
					int num4 = num2;
					Range range;
					Range range2;
					for (int j = 1; j <= num4; range = null, range2 = null, j++)
					{
						range = (Range)C.Cells[i, j];
						range2 = (Range)B.Cells[D - 1 + i, j];
						if (Operators.ConditionalCompareObjectEqual(range.Formula, string.Empty, TextCompare: false))
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
							range2.Clear();
							continue;
						}
						if (Conversions.ToBoolean(range.HasFormula))
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
							string numberFormat = Conversions.ToString(range.NumberFormat);
							string formula = range.Formula.ToString();
							Range range3 = ((_Worksheet)C.Worksheet).get_Range((object)VH.A(60932), RuntimeHelpers.GetObjectValue(Missing.Value));
							range3.Formula = formula;
							range3.Cut(range2);
							_ = null;
							Range obj = (Range)B.Cells[D - 1 + i, j];
							obj.Font.Color = F;
							obj.NumberFormat = numberFormat;
							_ = null;
							continue;
						}
						if (!Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range.Value2)))
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
							if (!KH.A.AutoColorText)
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
						}
						range2.Font.Color = E;
					}
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
				C = ((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[D, 1]), RuntimeHelpers.GetObjectValue(B.Cells[Operators.SubtractObject(Operators.AddObject(D, C.Rows.CountLarge), 1), RuntimeHelpers.GetObjectValue(C.Columns.CountLarge)]));
			}
			else
			{
				C.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
				Range obj2 = (Range)B.Cells[D, 1];
				obj2.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				obj2.Select();
				B.Paste(RuntimeHelpers.GetObjectValue(Missing.Value), true);
				((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[D, 1]), RuntimeHelpers.GetObjectValue(B.Cells[D + num, num2])).Font.Color = F;
				C = (Range)application.Selection;
			}
			application.CutCopyMode = (XlCutCopyMode)0;
			application = null;
		}
	}

	private static NE A(Worksheet A, ChartObject B)
	{
		int column = B.BottomRightCell.Column;
		int num = 4;
		NE result = default(NE);
		checked
		{
			int num2 = column + 1;
			Range range = ((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[2, num2]), RuntimeHelpers.GetObjectValue(A.Cells[2, column + 2]));
			range.Interior.Color = QuickCharts2.OPTIONS_DARK_YELLOW;
			range.VerticalAlignment = XlVAlign.xlVAlignCenter;
			range.RowHeight = 22;
			Range obj = (Range)range.Cells[1, 1];
			obj.Value2 = VH.A(60947);
			obj.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.White);
			obj.Font.Size = 14;
			_ = null;
			_ = null;
			((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[num - 1, num2]), RuntimeHelpers.GetObjectValue(A.Cells[E, column + 2])).Interior.Color = QuickCharts2.OPTIONS_TABLE_FILL;
			QuickCharts2.FormatOptionsHeader((Range)A.Cells[num, num2], VH.A(81601));
			Range range2 = (Range)A.Cells[num + 1, num2];
			Range range3 = range2;
			range3.Value2 = true;
			try
			{
				range3.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			CheckBox obj2 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range3.Left, range3.Top, range3.Width, range3.Height }, null, null, null);
			obj2.Text = VH.A(81622);
			obj2.LinkedCell = range3.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj2.Value = Microsoft.Office.Interop.Excel.Constants.xlBoth;
			range3 = null;
			result.A = range2;
			range2 = (Range)A.Cells[num + 2, num2];
			Range range4 = range2;
			range4.Value2 = false;
			try
			{
				range4.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			CheckBox obj3 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range4.Left, range4.Top, range4.Width, range4.Height }, null, null, null);
			obj3.Text = VH.A(81647);
			obj3.LinkedCell = range4.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj3.Value = Microsoft.Office.Interop.Excel.Constants.xlOff;
			range4 = null;
			result.B = range2;
			range2 = (Range)A.Cells[num + 3, num2];
			Range range5 = range2;
			range5.Value2 = false;
			try
			{
				range5.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			CheckBox obj4 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range5.Left, range5.Top, range5.Width, range5.Height }, null, null, null);
			obj4.Text = VH.A(81670);
			obj4.LinkedCell = range5.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj4.Value = Microsoft.Office.Interop.Excel.Constants.xlOff;
			range5 = null;
			result.C = range2;
			((Range)A.Cells[num + 5, num2]).Value2 = VH.A(81691);
			range2 = (Range)A.Cells[num + 5, num2 + 1];
			range2.Value2 = 0;
			range2.Validation.Add(XlDVType.xlValidateInputOnly, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			Validation validation = range2.Validation;
			validation.InputMessage = VH.A(81724);
			validation.ShowError = false;
			_ = null;
			QuickCharts2.FormatOptionsInput(range2);
			result.H = range2;
			QuickCharts2.FormatOptionsHeader((Range)A.Cells[num + 7, num2], VH.A(76781));
			((Range)A.Cells[num + 8, num2]).Value2 = VH.A(81968);
			((Range)A.Cells[num + 9, num2]).Value2 = VH.A(81995);
			((Range)A.Cells[num + 10, num2]).Value2 = VH.A(82020);
			num2 = column + 2;
			((Range)A.Columns[num2, RuntimeHelpers.GetObjectValue(Missing.Value)]).ColumnWidth = 18;
			range2 = (Range)A.Cells[num + 1, num2];
			Range range6 = range2;
			range6.Value2 = false;
			try
			{
				range6.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
			CheckBox obj5 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range6.Left, range6.Top, range6.Width, range6.Height }, null, null, null);
			obj5.Text = VH.A(82043);
			obj5.LinkedCell = range6.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj5.Value = Microsoft.Office.Interop.Excel.Constants.xlOff;
			range6 = null;
			result.D = range2;
			range2 = (Range)A.Cells[num + 8, num2];
			range2.Value2 = VH.A(82070);
			QuickCharts2.AddNumFormatValidation(range2);
			QuickCharts2.FormatOptionsInput(range2);
			result.E = range2;
			range2 = (Range)A.Cells[num + 9, num2];
			range2.Value2 = VH.A(82101);
			QuickCharts2.AddNumFormatValidation(range2);
			QuickCharts2.FormatOptionsInput(range2);
			result.F = range2;
			range2 = (Range)A.Cells[num + 10, num2];
			range2.Value2 = VH.A(82130);
			QuickCharts2.AddNumFormatValidation(range2);
			QuickCharts2.FormatOptionsInput(range2);
			result.G = range2;
			Border border = ((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[E, column + 1]), RuntimeHelpers.GetObjectValue(A.Cells[E, column + 2])).Borders[XlBordersIndex.xlEdgeBottom];
			border.LineStyle = XlLineStyle.xlContinuous;
			border.Weight = 2;
			border.Color = QuickCharts2.OPTIONS_DARK_YELLOW;
			_ = null;
			_ = null;
			range2 = null;
			return result;
		}
	}

	private static void A(Chart A, Range B, Range C, ME D, int E)
	{
		_ = A.Application;
		A.ChartType = XlChartType.xlBarStacked;
		string xValues = VH.A(48936) + ((Range)C.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
		Series series = ((SeriesCollection)A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(B.Columns[FootballField.m_A, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
		series.XValues = xValues;
		series.HasDataLabels = true;
		series.Format.Fill.Visible = MsoTriState.msoFalse;
		series.Format.Line.Visible = MsoTriState.msoFalse;
		((Microsoft.Office.Interop.Excel.DataLabels)series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Position = XlDataLabelPosition.xlLabelPositionInsideEnd;
		_ = null;
		Series series2 = ((SeriesCollection)A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(B.Columns[FootballField.C, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
		series2.XValues = xValues;
		series2.HasDataLabels = false;
		if (KH.A.ChartSeriesColors.Any())
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
			series2.Format.Fill.ForeColor.RGB = clsColors.RGB2Ole(KH.A.ChartSeriesColors[0]);
		}
		series2 = null;
		Series series3 = ((SeriesCollection)A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(B.Columns[FootballField.m_B, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
		series3.XValues = xValues;
		series3.HasDataLabels = true;
		series3.Format.Fill.Visible = MsoTriState.msoFalse;
		series3.Format.Line.Visible = MsoTriState.msoFalse;
		((Microsoft.Office.Interop.Excel.DataLabels)series3.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Position = XlDataLabelPosition.xlLabelPositionInsideBase;
		_ = null;
		Axis obj = (Axis)A.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue);
		obj.MinimumScaleIsAuto = false;
		obj.MaximumScale = obj.MaximumScale;
		_ = null;
		Axis obj2 = (Axis)A.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory);
		obj2.ReversePlotOrder = true;
		obj2.Crosses = XlAxisCrosses.xlAxisCrossesMaximum;
		_ = null;
		Series series4 = ((SeriesCollection)A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(B.Columns[5, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
		series4.ChartType = XlChartType.xlXYScatterLinesNoMarkers;
		series4.Values = VH.A(72891);
		series4.XValues = ((Range)B.Columns[5, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)2, RuntimeHelpers.GetObjectValue(Missing.Value));
		_ = null;
		FootballField.A(series4, D);
		FootballField.A(series4, D, ((Range)B.Columns[5, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)3).get_Resize((object)2, RuntimeHelpers.GetObjectValue(Missing.Value)), XlDataLabelPosition.xlLabelPositionRight);
		Series series5 = ((SeriesCollection)A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(B.Columns[6, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
		series5.ChartType = XlChartType.xlXYScatterLinesNoMarkers;
		series5.Values = VH.A(72891);
		series5.XValues = ((Range)B.Columns[6, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)2, RuntimeHelpers.GetObjectValue(Missing.Value));
		_ = null;
		FootballField.A(series5, D);
		FootballField.A(series5, D, ((Range)B.Columns[6, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)3).get_Resize((object)2, RuntimeHelpers.GetObjectValue(Missing.Value)), XlDataLabelPosition.xlLabelPositionRight);
		Series series6 = ((SeriesCollection)A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(B.Columns[7, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
		series6.ChartType = XlChartType.xlXYScatterLinesNoMarkers;
		series6.Values = VH.A(72891);
		series6.XValues = ((Range)B.Columns[7, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)2, RuntimeHelpers.GetObjectValue(Missing.Value));
		_ = null;
		FootballField.A(series6, D);
		FootballField.A(series6, D, ((Range)B.Columns[7, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)3).get_Resize((object)2, RuntimeHelpers.GetObjectValue(Missing.Value)), XlDataLabelPosition.xlLabelPositionRight);
		Axis obj3 = (Axis)A.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue, XlAxisGroup.xlSecondary);
		obj3.MaximumScale = 1.0;
		obj3.MinimumScale = 0.0;
		_ = null;
		((_Chart)A).set_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlValue, (object)XlAxisGroup.xlSecondary, (object)false);
	}

	private static void B(Chart A, Range B, Range C, ME D, int E)
	{
		_ = A.Application;
		int rGB = 0;
		A.ChartType = XlChartType.xlColumnStacked;
		string b = VH.A(48936) + ((Range)C.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
		string c = VH.A(48936) + ((Range)C.Columns[5, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
		Series series = ((SeriesCollection)A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(B.Columns[FootballField.m_B, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
		FootballField.A(series, b);
		((Microsoft.Office.Interop.Excel.DataLabels)series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Position = XlDataLabelPosition.xlLabelPositionAbove;
		Series series2 = ((SeriesCollection)A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(B.Columns[FootballField.m_A, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
		FootballField.A(series2, b);
		((Microsoft.Office.Interop.Excel.DataLabels)series2.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Position = XlDataLabelPosition.xlLabelPositionBelow;
		Series a = ((SeriesCollection)A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(B.Columns[5, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
		FootballField.A(a, D, c);
		FootballField.A(a, D, ((Range)B.Columns[5, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)3).get_Resize(Operators.SubtractObject(B.Rows.CountLarge, 1), RuntimeHelpers.GetObjectValue(Missing.Value)), XlDataLabelPosition.xlLabelPositionAbove);
		Series a2 = ((SeriesCollection)A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(B.Columns[6, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
		FootballField.A(a2, D, c);
		FootballField.A(a2, D, ((Range)B.Columns[6, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)3).get_Resize(Operators.SubtractObject(B.Rows.CountLarge, 1), RuntimeHelpers.GetObjectValue(Missing.Value)), XlDataLabelPosition.xlLabelPositionAbove);
		Series a3 = ((SeriesCollection)A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(B.Columns[7, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
		FootballField.A(a3, D, c);
		FootballField.A(a3, D, ((Range)B.Columns[7, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)3).get_Resize(Operators.SubtractObject(B.Rows.CountLarge, 1), RuntimeHelpers.GetObjectValue(Missing.Value)), XlDataLabelPosition.xlLabelPositionAbove);
		try
		{
			if (KH.A.ChartSeriesColors.Any())
			{
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
					rGB = clsColors.RGB2Ole(KH.A.ChartSeriesColors[0]);
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
		ChartGroup obj = (ChartGroup)A.LineGroups(1);
		obj.HasUpDownBars = true;
		obj.GapWidth = D.A;
		ChartFormat format = obj.UpBars.Format;
		format.Fill.ForeColor.RGB = rGB;
		format.Line.Visible = MsoTriState.msoFalse;
		_ = null;
		ChartFormat format2 = obj.DownBars.Format;
		format2.Fill.ForeColor.RGB = rGB;
		format2.Line.Visible = MsoTriState.msoFalse;
		_ = null;
		_ = null;
	}

	private static void A(Series A, string B)
	{
		A.ChartType = XlChartType.xlLine;
		A.MarkerStyle = XlMarkerStyle.xlMarkerStyleNone;
		A.Format.Line.Visible = MsoTriState.msoFalse;
		A.XValues = B;
		A.HasDataLabels = true;
		_ = null;
	}

	private static void A(Series A, ME B)
	{
		LineFormat line = A.Format.Line;
		line.ForeColor.RGB = B.B;
		line.DashStyle = B.A;
		_ = null;
	}

	private static void A(Series A, ME B, string C)
	{
		A.ChartType = XlChartType.xlXYScatterLinesNoMarkers;
		A.AxisGroup = XlAxisGroup.xlPrimary;
		A.XValues = C;
		A.Format.Line.Visible = MsoTriState.msoFalse;
		A.HasErrorBars = true;
		A.ErrorBars.EndStyle = XlEndStyleCap.xlNoCap;
		A.ErrorBar(XlErrorBarDirection.xlX, XlErrorBarInclude.xlErrorBarIncludePlusValues, XlErrorBarType.xlErrorBarTypeFixedValue, 1, RuntimeHelpers.GetObjectValue(Missing.Value));
		A.ErrorBar(XlErrorBarDirection.xlY, XlErrorBarInclude.xlErrorBarIncludeNone, XlErrorBarType.xlErrorBarTypeFixedValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		LineFormat line = A.ErrorBars.Format.Line;
		line.Weight = 1.5f;
		line.ForeColor.RGB = B.B;
		line.DashStyle = B.A;
		_ = null;
		_ = null;
	}

	private static void A(Series A, ME B, Range C, XlDataLabelPosition D)
	{
		A.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		A.HasLeaderLines = false;
		Microsoft.Office.Interop.Excel.Application application = C.Application;
		if (Conversion.Val(application.Version) < 15.0 && B.A == LE.B)
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
			application.ScreenUpdating = true;
			application.ScreenUpdating = false;
		}
		application = null;
		DataLabels.LinkToRange((Microsoft.Office.Interop.Excel.DataLabels)A.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), C);
		Microsoft.Office.Interop.Excel.DataLabels obj = (Microsoft.Office.Interop.Excel.DataLabels)A.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
		Font2 font = obj.Format.TextFrame2.TextRange.Font;
		font.Fill.ForeColor.RGB = B.B;
		font.Bold = MsoTriState.msoTrue;
		_ = null;
		obj.Position = D;
		_ = null;
		_ = null;
	}

	public static Range Example(Worksheet ws)
	{
		Worksheet worksheet = ws;
		((_Worksheet)worksheet).get_Range((object)VH.A(76929), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(82157);
		((_Worksheet)worksheet).get_Range((object)VH.A(76965), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(82172);
		((_Worksheet)worksheet).get_Range((object)VH.A(76981), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(82187);
		((_Worksheet)worksheet).get_Range((object)VH.A(76997), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(82202);
		((_Worksheet)worksheet).get_Range((object)VH.A(78364), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(82217);
		((_Worksheet)worksheet).get_Range((object)VH.A(78390), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(82232);
		((_Worksheet)worksheet).get_Range((object)VH.A(76877), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(81373);
		((_Worksheet)worksheet).get_Range((object)VH.A(76945), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 6;
		((_Worksheet)worksheet).get_Range((object)VH.A(61417), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 8;
		((_Worksheet)worksheet).get_Range((object)VH.A(61422), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 7;
		((_Worksheet)worksheet).get_Range((object)VH.A(61427), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 6;
		((_Worksheet)worksheet).get_Range((object)VH.A(61439), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 8.75;
		((_Worksheet)worksheet).get_Range((object)VH.A(61451), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 9;
		((_Worksheet)worksheet).get_Range((object)VH.A(57617), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(81380);
		((_Worksheet)worksheet).get_Range((object)VH.A(76950), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 10;
		((_Worksheet)worksheet).get_Range((object)VH.A(61486), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 12.75;
		((_Worksheet)worksheet).get_Range((object)VH.A(61502), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 8.5;
		((_Worksheet)worksheet).get_Range((object)VH.A(61507), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 10;
		((_Worksheet)worksheet).get_Range((object)VH.A(61512), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 12;
		((_Worksheet)worksheet).get_Range((object)VH.A(61517), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 11;
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
