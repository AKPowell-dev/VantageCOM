using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class ChartElements
{
	public static void Legend(Microsoft.Office.Interop.Word.Shape shp)
	{
		Chart chart = shp.Chart;
		if (chart.HasLegend)
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
			int num = Conversions.ToInteger(NewLateBinding.LateGet(chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, XC.A(22611), new object[0], null, null, null));
			int num2 = Conversions.ToInteger(NewLateBinding.LateGet(chart.Legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value)), null, XC.A(22611), new object[0], null, null, null));
			if (num != num2)
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
				Main.Analysis.Errors.Add(new ChartLegendEntryMissing(shp, string.Format(XC.A(22622), num, num2), chart.Legend));
			}
		}
		chart = null;
	}

	public static void MissingDataLabels(Microsoft.Office.Interop.Word.Shape shp)
	{
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = ((IEnumerable)shp.Chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
				IEnumerator enumerator2 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					IMsoSeries msoSeries = (IMsoSeries)enumerator.Current;
					if (!msoSeries.HasDataLabels)
					{
						continue;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					int num = 0;
					int num2 = 0;
					List<IMsoDataLabel> list = new List<IMsoDataLabel>();
					try
					{
						enumerator2 = ((IEnumerable)msoSeries.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						while (enumerator2.MoveNext())
						{
							IMsoDataLabel msoDataLabel = (IMsoDataLabel)enumerator2.Current;
							if (msoDataLabel.ShowValue)
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
								num2++;
								list.Add(msoDataLabel);
							}
							else
							{
								num++;
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_00c7;
							}
							continue;
							end_IL_00c7:
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
					if (num2 > 0)
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
						if (num > 0)
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
							Main.Analysis.Errors.Add(new ChartDataLabelMissing(shp, msoSeries, list));
						}
					}
					list = null;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						return;
					}
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
	}

	public static void DataLabelConsistency(Microsoft.Office.Interop.Word.Shape shp)
	{
		int num = 0;
		int num2 = 0;
		IEnumerator enumerator = ((IEnumerable)shp.Chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
		checked
		{
			try
			{
				while (enumerator.MoveNext())
				{
					if (((IMsoSeries)enumerator.Current).HasDataLabels)
					{
						num++;
					}
					else
					{
						num2++;
					}
					if (num <= 0)
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
					if (num2 <= 0)
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
						Main.Analysis.Errors.Add(new ChartDataLabelsInconsistent(shp));
						return;
					}
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						return;
					}
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
		}
	}

	public static void DataLabelFormat(Microsoft.Office.Interop.Word.Shape shp)
	{
		List<string> list = new List<string>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = ((IEnumerable)shp.Chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				IMsoSeries msoSeries = (IMsoSeries)enumerator.Current;
				if (!msoSeries.HasDataLabels)
				{
					continue;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				try
				{
					enumerator2 = ((IEnumerable)msoSeries.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
					while (enumerator2.MoveNext())
					{
						IMsoDataLabel msoDataLabel = (IMsoDataLabel)enumerator2.Current;
						if (!msoDataLabel.ShowValue)
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
							break;
						}
						list.Add(msoDataLabel.NumberFormat);
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_00be;
						}
						continue;
						end_IL_00be:
						break;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (3)
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
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_00f5;
				}
				continue;
				end_IL_00f5:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		IEnumerable<IGrouping<B<string>, string>> enumerable = from A in list
			group A by new B<string>(A.ToString());
		if (enumerable.Count() > 1)
		{
			List<string> list2 = new List<string>();
			List<string> list3 = new List<string>();
			IEnumerator<IGrouping<B<string>, string>> enumerator3 = default(IEnumerator<IGrouping<B<string>, string>>);
			try
			{
				enumerator3 = enumerable.GetEnumerator();
				while (enumerator3.MoveNext())
				{
					IGrouping<B<string>, string> current = enumerator3.Current;
					list3.Add(current.Key.Format + XC.A(22691) + current.Count() + XC.A(20696));
					list2.Add(current.Key.Format);
					current = null;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_01e6;
					}
					continue;
					end_IL_01e6:
					break;
				}
			}
			finally
			{
				if (enumerator3 != null)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						enumerator3.Dispose();
						break;
					}
				}
			}
			Main.Analysis.Errors.Add(new ChartDataLabelNumberFormats(shp, list3, string.Join(XC.A(22698), list3.ToArray()), list2));
			list2 = null;
			list3 = null;
		}
		enumerable = null;
	}

	public static void CheckDataLabelPosition(Microsoft.Office.Interop.Word.Shape shp)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = ((IEnumerable)shp.Chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				IMsoSeries msoSeries = (IMsoSeries)enumerator.Current;
				if (!msoSeries.HasDataLabels)
				{
					continue;
				}
				IMsoDataLabels msoDataLabels = (IMsoDataLabels)msoSeries.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
				int num = 0;
				try
				{
					try
					{
						enumerator2 = msoDataLabels.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							_ = (IMsoDataLabel)enumerator2.Current;
							_ = (Point)msoSeries.Points(num);
							num = checked(num + 1);
						}
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
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					return;
				}
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
}
