using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class PlotOrder
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<KeyValuePair<Series, double>, double> A;

		public static Func<KeyValuePair<Series, double>, Series> A;

		public static Func<KeyValuePair<Series, double>, double> B;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal double A(KeyValuePair<Series, double> A)
		{
			return A.Value;
		}

		[SpecialName]
		internal Series A(KeyValuePair<Series, double> A)
		{
			return A.Key;
		}

		[SpecialName]
		internal double B(KeyValuePair<Series, double> A)
		{
			return A.Value;
		}
	}

	public static void SmartSort()
	{
		if (!Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		checked
		{
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
				Chart chart = Helpers.SelectedChart();
				bool flag = false;
				if (chart != null)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
						{
							int count = ((ChartGroups)chart.ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value))).Count;
							Dictionary<Series, double> dictionary;
							Dictionary<Series, double> dictionary2;
							SeriesCollection seriesCollection;
							for (int i = 1; i <= count; i++)
							{
								try
								{
									dictionary = new Dictionary<Series, double>();
									seriesCollection = (SeriesCollection)((ChartGroups)chart.ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value))).Item(i).SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value));
									int count2 = seriesCollection.Count;
									for (int j = 1; j <= count2; j++)
									{
										Series series = seriesCollection.Item(j);
										switch (series.ChartType)
										{
										case XlChartType.xlColumnStacked:
										case XlChartType.xlColumnStacked100:
										case XlChartType.xl3DColumnStacked:
										case XlChartType.xl3DColumnStacked100:
										case XlChartType.xlBarStacked:
										case XlChartType.xlBarStacked100:
										case XlChartType.xl3DBarStacked:
										case XlChartType.xl3DBarStacked100:
											dictionary.Add(seriesCollection.Item(j), Conversions.ToDouble(NewLateBinding.LateIndexGet(series.Values, new object[1] { Information.UBound((Array)series.Values) }, null)));
											break;
										}
										series = null;
									}
									while (true)
									{
										switch (2)
										{
										case 0:
											break;
										default:
											if (dictionary.Any())
											{
												Dictionary<Series, double> source = dictionary;
												Func<KeyValuePair<Series, double>, double> keySelector;
												if (_Closure_0024__.A == null)
												{
													keySelector = (_Closure_0024__.A = [SpecialName] (KeyValuePair<Series, double> A) => A.Value);
												}
												else
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
													keySelector = _Closure_0024__.A;
												}
												IOrderedEnumerable<KeyValuePair<Series, double>> source2 = source.OrderByDescending(keySelector);
												Func<KeyValuePair<Series, double>, Series> keySelector2;
												if (_Closure_0024__.A == null)
												{
													keySelector2 = (_Closure_0024__.A = [SpecialName] (KeyValuePair<Series, double> A) => A.Key);
												}
												else
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
													keySelector2 = _Closure_0024__.A;
												}
												Func<KeyValuePair<Series, double>, double> elementSelector;
												if (_Closure_0024__.B == null)
												{
													elementSelector = (_Closure_0024__.B = [SpecialName] (KeyValuePair<Series, double> A) => A.Value);
												}
												else
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
													elementSelector = _Closure_0024__.B;
												}
												dictionary2 = source2.ToDictionary(keySelector2, elementSelector);
												int num = 1;
												foreach (KeyValuePair<Series, double> item in dictionary2)
												{
													item.Key.PlotOrder = num;
													num++;
												}
												flag = true;
											}
											goto end_IL_016f;
										}
										continue;
										end_IL_016f:
										break;
									}
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									Forms.ErrorMessage(ex2.Message);
									ProjectData.ClearProjectError();
								}
							}
							dictionary = null;
							dictionary2 = null;
							seriesCollection = null;
							chart = null;
							if (flag)
							{
								clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(65379));
							}
							else
							{
								Forms.WarningMessage(VH.A(65412));
							}
							return;
						}
						}
					}
				}
				Helpers.NoChartMessage();
				return;
			}
		}
	}
}
