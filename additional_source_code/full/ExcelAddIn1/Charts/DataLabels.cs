using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class DataLabels
{
	private enum RD
	{
		A,
		B
	}

	private enum SD
	{
		A = -4142,
		B = -4152,
		C = -4131,
		D = 0,
		E = 1,
		F = 2,
		G = 3
	}

	private struct TD
	{
		public SD A;

		public SD B;

		public SD C;

		public SD D;

		public RD A;

		public bool A;

		public bool B;

		public int A;
	}

	public static void ReplaceMissingLabels(string tag)
	{
		if (!A())
		{
			return;
		}
		if (Operators.CompareString(tag, VH.A(70109), TextCompare: false) != 0)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (Operators.CompareString(tag, VH.A(70118), TextCompare: false) != 0)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								if (Operators.CompareString(tag, VH.A(70127), TextCompare: false) != 0)
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											return;
										}
									}
								}
								A(A: false, B: true);
								return;
							}
						}
					}
					A(A: true, B: false);
					return;
				}
			}
		}
		A(A: true, B: true);
	}

	private static void A(bool A, bool B)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Axis axis = default(Axis);
		Chart chart = default(Chart);
		Axis axis2 = default(Axis);
		SeriesCollection seriesCollection = default(SeriesCollection);
		int count = default(int);
		int num5 = default(int);
		Series series = default(Series);
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 597:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_0011;
						case 4:
							goto IL_002c;
						case 5:
							goto IL_0034;
						case 6:
							goto IL_0051;
						case 7:
							goto IL_0066;
						case 8:
							goto IL_007a;
						case 9:
							goto IL_0087;
						case 10:
							goto IL_0092;
						case 11:
							goto IL_0100;
						case 12:
							goto IL_0103;
						case 13:
							goto IL_011f;
						case 14:
							goto IL_0125;
						case 15:
							goto IL_012e;
						case 16:
							goto IL_0166;
						case 17:
							goto IL_016d;
						case 18:
							goto IL_0185;
						case 19:
							goto IL_0190;
						case 20:
							goto IL_0193;
						case 21:
							goto IL_01ab;
						case 22:
							goto IL_01cd;
						case 24:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 23:
						case 25:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0190:
					axis = null;
					goto IL_0193;
					IL_0007:
					num2 = 2;
					chart = Helpers.SelectedChart();
					goto IL_0011;
					IL_0011:
					num2 = 3;
					if (chart == null)
					{
						break;
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
					goto IL_002c;
					IL_0193:
					num2 = 20;
					goto IL_0196;
					IL_0185:
					num2 = 18;
					axis.HasDisplayUnitLabel = true;
					goto IL_0190;
					IL_0166:
					num2 = 16;
					axis = axis2;
					goto IL_016d;
					IL_002c:
					num2 = 4;
					if (A)
					{
						goto IL_0034;
					}
					goto IL_0125;
					IL_0034:
					num2 = 5;
					seriesCollection = (SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value));
					goto IL_0051;
					IL_0051:
					num2 = 6;
					count = seriesCollection.Count;
					num5 = 1;
					goto IL_010c;
					IL_010c:
					if (num5 <= count)
					{
						goto IL_0066;
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
					goto IL_011f;
					IL_016d:
					num2 = 17;
					if (!axis.HasDisplayUnitLabel)
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
						goto IL_0185;
					}
					goto IL_0190;
					IL_011f:
					num2 = 13;
					seriesCollection = null;
					goto IL_0125;
					IL_0066:
					num2 = 7;
					series = seriesCollection.Item(num5);
					goto IL_007a;
					IL_007a:
					num2 = 8;
					if (!series.HasDataLabels)
					{
						goto IL_0087;
					}
					goto IL_0100;
					IL_0087:
					num2 = 9;
					series.HasDataLabels = true;
					goto IL_0092;
					IL_0092:
					num2 = 10;
					series.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					goto IL_0100;
					IL_0100:
					series = null;
					goto IL_0103;
					IL_0103:
					num2 = 12;
					num5 = checked(num5 + 1);
					goto IL_010c;
					IL_0125:
					num2 = 14;
					if (B)
					{
						goto IL_012e;
					}
					goto IL_01cd;
					IL_012e:
					num2 = 15;
					enumerator = ((IEnumerable)chart.Axes(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
					goto IL_0196;
					IL_0196:
					if (enumerator.MoveNext())
					{
						axis2 = (Axis)enumerator.Current;
						goto IL_0166;
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
					goto IL_01ab;
					IL_01cd:
					num2 = 22;
					chart = null;
					goto end_IL_0000_3;
					IL_01ab:
					num2 = 21;
					if (enumerator is IDisposable)
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
						(enumerator as IDisposable).Dispose();
					}
					goto IL_01cd;
					end_IL_0000_2:
					break;
				}
				num2 = 24;
				Helpers.NoChartMessage();
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 597;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
		{
			return;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static void LinkFormatsToCells()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		Axis axis = default(Axis);
		Series series = default(Series);
		Chart chart = default(Chart);
		SeriesCollection seriesCollection = default(SeriesCollection);
		int count = default(int);
		int num5 = default(int);
		IEnumerator enumerator = default(IEnumerator);
		Microsoft.Office.Interop.Excel.DataLabels dataLabels = default(Microsoft.Office.Interop.Excel.DataLabels);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					if (!A())
					{
						goto end_IL_0000;
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
					goto IL_0021;
				case 560:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000_2;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 3:
							goto IL_0021;
						case 4:
							goto IL_0028;
						case 5:
							goto IL_0030;
						case 6:
							goto IL_0038;
						case 7:
							goto IL_0055;
						case 8:
							goto IL_0067;
						case 9:
							goto IL_007b;
						case 10:
							goto IL_0091;
						case 11:
							goto IL_00ae;
						case 12:
							goto IL_00b9;
						case 13:
							goto IL_00bc;
						case 14:
							goto IL_00d5;
						case 15:
							goto IL_0109;
						case 16:
							goto IL_0121;
						case 17:
							goto IL_0133;
						case 18:
							goto IL_0141;
						case 19:
							goto IL_0163;
						case 20:
							goto IL_0184;
						case 21:
							goto IL_0198;
						case 22:
							goto IL_019e;
						case 23:
							goto IL_01a3;
						case 25:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 24:
						case 26:
							goto end_IL_0000;
						}
						goto default;
					}
					IL_0109:
					num2 = 15;
					if (axis.HasDisplayUnitLabel)
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
						goto IL_0121;
					}
					goto IL_0133;
					IL_00b9:
					series = null;
					goto IL_00bc;
					IL_0121:
					num2 = 16;
					axis.TickLabels.NumberFormatLinked = true;
					goto IL_0133;
					IL_0133:
					num2 = 17;
					goto IL_0136;
					IL_0021:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0028;
					IL_0028:
					num2 = 4;
					chart = Helpers.SelectedChart();
					goto IL_0030;
					IL_0030:
					num2 = 5;
					if (chart == null)
					{
						break;
					}
					goto IL_0038;
					IL_0038:
					num2 = 6;
					seriesCollection = (SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value));
					goto IL_0055;
					IL_0055:
					num2 = 7;
					count = seriesCollection.Count;
					num5 = 1;
					goto IL_00c5;
					IL_00c5:
					if (num5 <= count)
					{
						goto IL_0067;
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
					goto IL_00d5;
					IL_0067:
					num2 = 8;
					series = seriesCollection.Item(num5);
					goto IL_007b;
					IL_00d5:
					num2 = 14;
					enumerator = ((IEnumerable)chart.Axes(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
					goto IL_0136;
					IL_0136:
					if (enumerator.MoveNext())
					{
						axis = (Axis)enumerator.Current;
						goto IL_0109;
					}
					goto IL_0141;
					IL_0141:
					num2 = 18;
					if (enumerator is IDisposable)
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
						(enumerator as IDisposable).Dispose();
					}
					goto IL_0163;
					IL_00bc:
					num2 = 13;
					num5 = checked(num5 + 1);
					goto IL_00c5;
					IL_0091:
					num2 = 10;
					dataLabels = (Microsoft.Office.Interop.Excel.DataLabels)series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
					goto IL_00ae;
					IL_0163:
					num2 = 19;
					if (chart.Application.Calculation == XlCalculation.xlCalculationManual)
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
						goto IL_0184;
					}
					goto IL_0198;
					IL_00ae:
					num2 = 11;
					dataLabels.NumberFormatLinked = true;
					goto IL_00b9;
					IL_0184:
					num2 = 20;
					Forms.InfoMessage(VH.A(70136));
					goto IL_0198;
					IL_0198:
					num2 = 21;
					dataLabels = null;
					goto IL_019e;
					IL_019e:
					num2 = 22;
					chart = null;
					goto IL_01a3;
					IL_01a3:
					num2 = 23;
					seriesCollection = null;
					goto end_IL_0000;
					IL_007b:
					num2 = 9;
					if (series.HasDataLabels)
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
						goto IL_0091;
					}
					goto IL_00b9;
					end_IL_0000_3:
					break;
				}
				num2 = 25;
				Helpers.NoChartMessage();
				break;
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 560;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
	}

	public static void LinkToRange(Microsoft.Office.Interop.Excel.DataLabels labels, Range rng)
	{
		labels.Format.TextFrame2.TextRange.InsertChartField(MsoChartFieldType.msoChartFieldRange, VH.A(48936) + rng.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), 0);
		labels.ShowRange = true;
		labels.ShowValue = false;
		_ = null;
	}

	public static void AttachLabelsToPoints()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		Chart chart = default(Chart);
		Series series = default(Series);
		Series series2 = default(Series);
		Application application = default(Application);
		string text = default(string);
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					if (!A())
					{
						goto end_IL_0000;
					}
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
					goto IL_0021;
				case 729:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000_2;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 3:
							goto IL_0021;
						case 4:
							goto IL_0028;
						case 5:
							goto IL_0032;
						case 6:
							goto IL_003a;
						case 7:
							goto IL_0046;
						case 8:
							goto IL_0050;
						case 9:
							goto IL_0084;
						case 10:
							goto IL_00d8;
						case 13:
							goto IL_0114;
						case 11:
						case 12:
						case 14:
							goto IL_0123;
						case 15:
							goto IL_014e;
						case 16:
							goto IL_0185;
						case 17:
							goto IL_018c;
						case 18:
							goto IL_0197;
						case 19:
							goto IL_01e1;
						case 20:
							goto IL_01e4;
						case 21:
							goto IL_01fc;
						case 22:
							goto IL_021e;
						case 23:
							goto IL_0229;
						case 24:
							goto IL_022f;
						case 25:
							goto IL_0234;
						case 27:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 26:
						case 28:
							goto end_IL_0000;
						}
						goto default;
					}
					IL_022f:
					num2 = 24;
					chart = null;
					goto IL_0234;
					IL_0234:
					num2 = 25;
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(70308));
					goto end_IL_0000;
					IL_01e1:
					series = null;
					goto IL_01e4;
					IL_0185:
					num2 = 16;
					series = series2;
					goto IL_018c;
					IL_0021:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0028;
					IL_0028:
					num2 = 4;
					chart = Helpers.SelectedChart();
					goto IL_0032;
					IL_0032:
					num2 = 5;
					if (chart == null)
					{
						break;
					}
					goto IL_003a;
					IL_003a:
					num2 = 6;
					application = chart.Application;
					goto IL_0046;
					IL_0046:
					num2 = 7;
					application.ScreenUpdating = false;
					goto IL_0050;
					IL_0050:
					num2 = 8;
					text = Conversions.ToString(NewLateBinding.LateGet(chart.SeriesCollection(1), null, VH.A(68956), new object[0], null, null, null));
					goto IL_0084;
					IL_0084:
					num2 = 9;
					text = Strings.Mid(text, Strings.InStr(Strings.InStr(text, VH.A(2378)), text, Strings.Mid(Strings.Left(text, checked(Strings.InStr(text, VH.A(7827)) - 1)), 9)));
					goto IL_00d8;
					IL_00d8:
					num2 = 10;
					text = Strings.Left(text, checked(Strings.InStr(Strings.InStr(text, VH.A(7827)), text, VH.A(2378)) - 1));
					goto IL_0123;
					IL_0123:
					num2 = 12;
					if (Operators.CompareString(Strings.Left(text, 1), VH.A(2378), TextCompare: false) == 0)
					{
						goto IL_0114;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					goto IL_014e;
					IL_018c:
					num2 = 17;
					series.HasDataLabels = true;
					goto IL_0197;
					IL_014e:
					num2 = 15;
					enumerator = ((IEnumerable)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
					goto IL_01e7;
					IL_01e7:
					if (enumerator.MoveNext())
					{
						series2 = (Series)enumerator.Current;
						goto IL_0185;
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
					goto IL_01fc;
					IL_0197:
					num2 = 18;
					LinkToRange((Microsoft.Office.Interop.Excel.DataLabels)series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), ((_Application)application).get_Range((object)text, RuntimeHelpers.GetObjectValue(Missing.Value)).get_Offset((object)0, (object)(-1)));
					goto IL_01e1;
					IL_01fc:
					num2 = 21;
					if (enumerator is IDisposable)
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
						(enumerator as IDisposable).Dispose();
					}
					goto IL_021e;
					IL_01e4:
					num2 = 20;
					goto IL_01e7;
					IL_0114:
					num2 = 13;
					text = Strings.Mid(text, 2);
					goto IL_0123;
					IL_021e:
					num2 = 22;
					application.ScreenUpdating = true;
					goto IL_0229;
					IL_0229:
					num2 = 23;
					application = null;
					goto IL_022f;
					end_IL_0000_3:
					break;
				}
				num2 = 27;
				Helpers.NoChartMessage();
				break;
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 729;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000:
			break;
		}
		if (num == 0)
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static void LabelPoints()
	{
		if (!A())
		{
			return;
		}
		bool flag = true;
		try
		{
			Chart chart = Helpers.SelectedChart();
			TD c = default(TD);
			int num2 = default(int);
			if (chart != null)
			{
				while (true)
				{
					wpfLabelPoints wpfLabelPoints2;
					XmlDocument settingsXml;
					XmlNode documentElement;
					switch (6)
					{
					case 0:
						break;
					default:
						{
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							settingsXml = KH.A.SettingsXml;
							documentElement = settingsXml.DocumentElement;
							wpfLabelPoints2 = new wpfLabelPoints();
							wpfLabelPoints2.cbxLast.SelectedIndex = Conversions.ToInteger(documentElement.SelectSingleNode(VH.A(70345)).InnerText);
							wpfLabelPoints2.cbxFirst.SelectedIndex = Conversions.ToInteger(documentElement.SelectSingleNode(VH.A(70394)).InnerText);
							wpfLabelPoints2.cbxMaximum.SelectedIndex = Conversions.ToInteger(documentElement.SelectSingleNode(VH.A(70445)).InnerText);
							wpfLabelPoints2.cbxMinimum.SelectedIndex = Conversions.ToInteger(documentElement.SelectSingleNode(VH.A(70492)).InnerText);
							if (Conversions.ToInteger(documentElement.SelectSingleNode(VH.A(70539)).InnerText) == 0)
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
								wpfLabelPoints2.radValue.IsChecked = true;
							}
							else
							{
								wpfLabelPoints2.radSeriesName.IsChecked = true;
							}
							wpfLabelPoints2.chkColor.IsChecked = Conversions.ToBoolean(documentElement.SelectSingleNode(VH.A(70582)).InnerText);
							wpfLabelPoints2.chkBold.IsChecked = Conversions.ToBoolean(documentElement.SelectSingleNode(VH.A(70627)).InnerText);
							wpfLabelPoints2.cbxRotation.SelectedIndex = Conversions.ToInteger(documentElement.SelectSingleNode(VH.A(70668)).InnerText);
							wpfLabelPoints2.ShowDialog();
							if (wpfLabelPoints2.DialogResult.HasValue)
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
								if (wpfLabelPoints2.DialogResult.Value)
								{
									c = default(TD);
									switch (wpfLabelPoints2.cbxLast.SelectedIndex)
									{
									case 0:
										c.A = SD.A;
										break;
									case 1:
										c.A = SD.B;
										break;
									case 2:
										c.A = SD.D;
										break;
									case 3:
										c.A = SD.E;
										break;
									case 4:
										c.A = SD.F;
										break;
									case 5:
										c.A = SD.G;
										break;
									}
									switch (wpfLabelPoints2.cbxFirst.SelectedIndex)
									{
									case 0:
										c.B = SD.A;
										break;
									case 1:
										c.B = SD.C;
										break;
									case 2:
										c.B = SD.D;
										break;
									case 3:
										c.B = SD.E;
										break;
									case 4:
										c.B = SD.F;
										break;
									case 5:
										c.B = SD.G;
										break;
									}
									switch (wpfLabelPoints2.cbxMaximum.SelectedIndex)
									{
									case 0:
										c.D = SD.A;
										break;
									case 1:
										c.D = SD.D;
										break;
									case 2:
										c.D = SD.F;
										break;
									case 3:
										c.D = SD.G;
										break;
									}
									switch (wpfLabelPoints2.cbxMinimum.SelectedIndex)
									{
									case 0:
										c.C = SD.A;
										break;
									case 1:
										c.C = SD.E;
										break;
									case 2:
										c.C = SD.F;
										break;
									case 3:
										c.C = SD.G;
										break;
									}
									switch (wpfLabelPoints2.cbxRotation.SelectedIndex)
									{
									case 0:
										c.A = 0;
										break;
									case 1:
										c.A = 90;
										break;
									case 2:
										c.A = -90;
										break;
									}
									if (wpfLabelPoints2.radValue.IsChecked == true)
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
										c.A = RD.A;
									}
									else
									{
										c.A = RD.B;
									}
									c.A = wpfLabelPoints2.chkColor.IsChecked.Value;
									c.B = wpfLabelPoints2.chkBold.IsChecked.Value;
									documentElement.SelectSingleNode(VH.A(70345)).InnerText = wpfLabelPoints2.cbxLast.SelectedIndex.ToString();
									documentElement.SelectSingleNode(VH.A(70394)).InnerText = wpfLabelPoints2.cbxFirst.SelectedIndex.ToString();
									documentElement.SelectSingleNode(VH.A(70492)).InnerText = wpfLabelPoints2.cbxMinimum.SelectedIndex.ToString();
									documentElement.SelectSingleNode(VH.A(70445)).InnerText = wpfLabelPoints2.cbxMaximum.SelectedIndex.ToString();
									documentElement.SelectSingleNode(VH.A(70668)).InnerText = wpfLabelPoints2.cbxRotation.SelectedIndex.ToString();
									XmlNode xmlNode = documentElement.SelectSingleNode(VH.A(70539));
									int a = (int)c.A;
									xmlNode.InnerText = a.ToString();
									documentElement.SelectSingleNode(VH.A(70582)).InnerText = (0 - (c.A ? 1 : 0)).ToString();
									documentElement.SelectSingleNode(VH.A(70627)).InnerText = (0 - (c.B ? 1 : 0)).ToString();
									KH.A.SaveSettings(settingsXml);
									goto IL_0612;
								}
							}
							flag = false;
							goto IL_0612;
						}
						IL_0612:
						wpfLabelPoints2 = null;
						settingsXml = null;
						documentElement = null;
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
							SeriesCollection seriesCollection = (SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value));
							int a = seriesCollection.Count;
							int num = 1;
							checked
							{
								while (true)
								{
									if (num > a)
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
										break;
									}
									Series series = seriesCollection.Item(num);
									XlChartType chartType = series.ChartType;
									if (chartType <= XlChartType.xl3DLine)
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
										if (chartType != XlChartType.xlXYScatter)
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
											if (chartType != XlChartType.xl3DLine)
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
												goto IL_07f1;
											}
										}
									}
									else if (chartType != XlChartType.xlLine)
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
										switch (chartType)
										{
										case XlChartType.xlColumnClustered:
										case XlChartType.xl3DColumnClustered:
										case XlChartType.xlBarClustered:
										case XlChartType.xl3DBarClustered:
										case XlChartType.xlLineStacked:
										case XlChartType.xlLineStacked100:
										case XlChartType.xlLineMarkers:
										case XlChartType.xlLineMarkersStacked:
										case XlChartType.xlLineMarkersStacked100:
										case XlChartType.xlXYScatterSmooth:
										case XlChartType.xlXYScatterSmoothNoMarkers:
										case XlChartType.xlXYScatterLines:
										case XlChartType.xlXYScatterLinesNoMarkers:
											break;
										default:
											goto IL_07f1;
										}
									}
									series.HasDataLabels = false;
									if (c.A == RD.B)
									{
										chart.HasLegend = false;
									}
									int count = ((Points)series.Points(RuntimeHelpers.GetObjectValue(Missing.Value))).Count;
									int e = 1;
									int f = count;
									if (count > 1095)
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
										if (c.D == SD.A)
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
											if (c.C == SD.A)
											{
												goto IL_07d1;
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
										}
										Forms.WarningMessage(VH.A(70709));
										num2++;
										break;
									}
									goto IL_07d1;
									IL_07f1:
									series = null;
									num++;
									continue;
									IL_07d1:
									A(chart, seriesCollection.Item(num), c, count, e, f);
									num2++;
									goto IL_07f1;
								}
								seriesCollection = null;
								if (num2 == 0)
								{
									Forms.WarningMessage(VH.A(70790));
								}
								clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(70891));
							}
						}
						chart = null;
						return;
					}
				}
			}
			Helpers.NoChartMessage();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(Chart A, Series B, TD C, int D, int E, int F)
	{
		Series series = B;
		checked
		{
			if (C.A != SD.A)
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
				DataLabels.A(B, D, C.A, C);
				F--;
			}
			if (C.B != SD.A)
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
				DataLabels.A(B, 1, C.B, C);
				E++;
			}
			if (C.C != SD.A)
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
				if (C.D != SD.A)
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
					double num = A.Application.WorksheetFunction.Min(RuntimeHelpers.GetObjectValue(series.Values), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					double num2 = A.Application.WorksheetFunction.Max(RuntimeHelpers.GetObjectValue(series.Values), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					int num3 = E;
					int num4 = F;
					for (int i = num3; i <= num4; i++)
					{
						if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateIndexGet(series.Values, new object[1] { i }, null), num, TextCompare: false))
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
							DataLabels.A(B, i, C.C, C);
						}
						else
						{
							if (!Operators.ConditionalCompareObjectEqual(NewLateBinding.LateIndexGet(series.Values, new object[1] { i }, null), num2, TextCompare: false))
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
							DataLabels.A(B, i, C.D, C);
						}
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
					goto IL_07fd;
				}
			}
			if (C.C != SD.A)
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
				double num = A.Application.WorksheetFunction.Min(RuntimeHelpers.GetObjectValue(series.Values), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				int num5 = E;
				int num6 = F;
				for (int j = num5; j <= num6; j++)
				{
					if (!Operators.ConditionalCompareObjectEqual(NewLateBinding.LateIndexGet(series.Values, new object[1] { j }, null), num, TextCompare: false))
					{
						continue;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
					DataLabels.A(B, j, C.C, C);
				}
			}
			else if (C.D != SD.A)
			{
				double num2 = A.Application.WorksheetFunction.Max(RuntimeHelpers.GetObjectValue(series.Values), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				int num7 = E;
				int num8 = F;
				for (int k = num7; k <= num8; k++)
				{
					if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateIndexGet(series.Values, new object[1] { k }, null), num2, TextCompare: false))
					{
						DataLabels.A(B, k, C.D, C);
					}
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					break;
				}
			}
			goto IL_07fd;
		}
		IL_07fd:
		series.HasLeaderLines = false;
		series = null;
	}

	private static void A(Series A, int B, SD C, TD D)
	{
		Point point = (Point)A.Points(B);
		point.HasDataLabel = true;
		point.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		if (D.A == RD.B)
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
			point.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			DataLabel dataLabel = ((Microsoft.Office.Interop.Excel.DataLabels)A.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Item(B);
			dataLabel.ShowSeriesName = true;
			dataLabel.ShowValue = false;
			_ = null;
		}
		point.DataLabel.NumberFormatLinked = true;
		try
		{
			point.DataLabel.Position = (XlDataLabelPosition)C;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.WarningMessage(VH.A(70916));
			ProjectData.ClearProjectError();
		}
		if (D.A)
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
			XlChartType chartType = A.ChartType;
			if (chartType <= XlChartType.xl3DColumnClustered)
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
				if (chartType != XlChartType.xlColumnClustered)
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
					if (chartType != XlChartType.xl3DColumnClustered)
					{
						goto IL_02bf;
					}
				}
			}
			else if (chartType != XlChartType.xlBarClustered && chartType != XlChartType.xl3DBarClustered)
			{
				goto IL_02bf;
			}
			try
			{
				if (point.Format.Fill.Visible == MsoTriState.msoTrue)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						point.DataLabel.Font.Color = point.Format.Fill.ForeColor.RGB;
						break;
					}
				}
				else if (point.Format.Line.Visible == MsoTriState.msoTrue)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						if (!(point.Format.Line.Weight > 0f))
						{
							break;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							point.DataLabel.Font.Color = point.Format.Line.ForeColor.RGB;
							break;
						}
						break;
					}
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		goto IL_02e9;
		IL_02e9:
		point.DataLabel.Font.Bold = D.B;
		point.DataLabel.Orientation = checked(-1 * D.A);
		point = null;
		return;
		IL_02bf:
		point.DataLabel.Font.Color = RuntimeHelpers.GetObjectValue(point.Border.Color);
		goto IL_02e9;
	}

	internal static bool A()
	{
		return Access.AllowExcelOperation((PlanType)4, (Restriction)1, false);
	}

	public static void RotateHorizontal()
	{
		A(XlOrientation.xlHorizontal);
	}

	public static void Rotate90()
	{
		A(XlOrientation.xlDownward);
	}

	public static void Rotate270()
	{
		A(XlOrientation.xlUpward);
	}

	public static void RotateStacked()
	{
		A(XlOrientation.xlVertical);
	}

	private static void A(XlOrientation A)
	{
		if (Helpers.SelectedChart() != null)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					object objectValue;
					try
					{
						objectValue = RuntimeHelpers.GetObjectValue(MH.A.Application.Selection);
						if (objectValue is Microsoft.Office.Interop.Excel.DataLabels)
						{
							((Microsoft.Office.Interop.Excel.DataLabels)objectValue).Orientation = A;
						}
						else if (objectValue is DataLabel)
						{
							((DataLabel)objectValue).Orientation = A;
						}
						clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(71102));
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					objectValue = null;
					return;
				}
				}
			}
		}
		Helpers.NoChartMessage();
	}
}
