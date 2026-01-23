using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using ExcelAddIn1.Charts;
using ExcelAddIn1.FastFormats.Charts.Objects;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.Config.Settings;
using MacabacusMacros.FastFormats.Charts;
using MacabacusMacros.ImportExport;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts;

public sealed class Apply
{
	public static void Selection()
	{
		if (!Access.AllowExcelOperation((PlanType)6, (Restriction)2, false))
		{
			return;
		}
		Application application = MH.A.Application;
		string text = string.Join(VH.A(75498), Constants.XML_FAST_FORMATS, Constants.XML_CHART_FORMATS, FormatConstants.NODE_CHART);
		try
		{
			XmlElement documentElement = KH.A.SettingsXml.DocumentElement;
			XmlNodeList xmlNodeList = documentElement.SelectNodes(text);
			if (xmlNodeList != null)
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
				if (xmlNodeList.Count != 0)
				{
					XmlNode xmlNode = documentElement.SelectSingleNode(text + VH.A(140758));
					if (xmlNode == null)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								Forms.WarningMessage(VH.A(140779));
								return;
							}
						}
					}
					documentElement = null;
					ExcelAddIn1.FastFormats.Charts.Objects.Chart b;
					try
					{
						b = new ExcelAddIn1.FastFormats.Charts.Objects.Chart(xmlNode);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						Forms.ErrorMessage(VH.A(140834));
						clsReporting.LogException(ex2);
						ProjectData.ClearProjectError();
						return;
					}
					(bool, List<Microsoft.Office.Interop.Excel.Chart>) tuple = Helpers.SelectedCharts(Apply.A);
					if (!tuple.Item1)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								return;
							}
						}
					}
					if (tuple.Item2 == null)
					{
						Forms.WarningMessage(VH.A(56295));
						return;
					}
					application.ScreenUpdating = false;
					try
					{
						using (List<Microsoft.Office.Interop.Excel.Chart>.Enumerator enumerator = tuple.Item2.GetEnumerator())
						{
							while (enumerator.MoveNext())
							{
								A(enumerator.Current, b, xmlNodeList);
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_019c;
								}
								continue;
								end_IL_019c:
								break;
							}
						}
						clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(140901));
						return;
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						Forms.ErrorMessage(ex4.Message);
						clsReporting.LogException(ex4);
						ProjectData.ClearProjectError();
						return;
					}
				}
			}
			Forms.WarningMessage(VH.A(140701));
		}
		finally
		{
			application.ScreenUpdating = true;
			(bool, List<Microsoft.Office.Interop.Excel.Chart>) tuple = default((bool, List<Microsoft.Office.Interop.Excel.Chart>));
			XmlNodeList xmlNodeList = null;
			application = null;
			ExcelAddIn1.FastFormats.Charts.Objects.Chart b = null;
			XmlNode xmlNode = null;
		}
	}

	private static Dictionary<XlChartType, List<Microsoft.Office.Interop.Excel.Series>> A(Microsoft.Office.Interop.Excel.Chart A)
	{
		Dictionary<XlChartType, List<Microsoft.Office.Interop.Excel.Series>> dictionary = new Dictionary<XlChartType, List<Microsoft.Office.Interop.Excel.Series>>();
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = ((IEnumerable)A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
				while (enumerator.MoveNext())
				{
					Microsoft.Office.Interop.Excel.Series series = (Microsoft.Office.Interop.Excel.Series)enumerator.Current;
					XlChartType chartCombinedType = ChartTypes.GetChartCombinedType(series.ChartType);
					if (dictionary.ContainsKey(chartCombinedType))
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
						dictionary[chartCombinedType].Add(series);
					}
					else
					{
						List<Microsoft.Office.Interop.Excel.Series> value = new List<Microsoft.Office.Interop.Excel.Series> { series };
						dictionary.Add(chartCombinedType, value);
					}
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_008f;
					}
					continue;
					end_IL_008f:
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
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return dictionary;
	}

	private static void A(Microsoft.Office.Interop.Excel.Chart A, ExcelAddIn1.FastFormats.Charts.Objects.Chart B, XmlNodeList C)
	{
		Dictionary<XlChartType, List<Microsoft.Office.Interop.Excel.Series>> dictionary = Apply.A(A);
		B.ApplyTo(A, dictionary);
		IEnumerator enumerator2 = default(IEnumerator);
		foreach (KeyValuePair<XlChartType, List<Microsoft.Office.Interop.Excel.Series>> item in dictionary)
		{
			try
			{
				enumerator2 = C.GetEnumerator();
				while (true)
				{
					if (enumerator2.MoveNext())
					{
						XmlNode xmlNode = (XmlNode)enumerator2.Current;
						string value = xmlNode.Attributes[FormatConstants.ATTR_TYPE].Value;
						if (string.IsNullOrEmpty(value))
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
						if (Conversions.ToInteger(value) != (int)item.Key)
						{
							continue;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							new ExcelAddIn1.FastFormats.Charts.Objects.Chart(xmlNode).ApplyTo(A, dictionary);
							break;
						}
						break;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_00b5;
						}
						continue;
						end_IL_00b5:
						break;
					}
					break;
				}
			}
			finally
			{
				if (enumerator2 is IDisposable)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						(enumerator2 as IDisposable).Dispose();
						break;
					}
				}
			}
		}
	}

	private static bool A(Microsoft.Office.Interop.Excel.Chart A)
	{
		if (clsImportExport.IsUnsupportedChartType(A))
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
					Forms.WarningMessage(VH.A(140936));
					return false;
				}
			}
		}
		if (ChartTypes.IsStockChart(A.ChartType))
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					Forms.WarningMessage(VH.A(141180));
					return false;
				}
			}
		}
		return true;
	}
}
