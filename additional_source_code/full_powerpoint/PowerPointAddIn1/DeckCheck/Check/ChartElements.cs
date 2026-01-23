using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class ChartElements
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<string, D<string>> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal D<string> A(string A)
		{
			return new D<string>(A.ToString());
		}
	}

	public static void Legend(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		Chart chart = shp.Chart;
		if (chart.HasLegend)
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
			int[] b = new int[11]
			{
				119, 117, 83, 84, 85, 86, -4120, 5, -4102, 68,
				71
			};
			int num = Conversions.ToInteger(NewLateBinding.LateGet(chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, AH.A(13955), new object[0], null, null, null));
			int num2 = Conversions.ToInteger(NewLateBinding.LateGet(chart.Legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value)), null, AH.A(13955), new object[0], null, null, null));
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
				if (clsCharts.A(shp, b))
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
					Main.Analysis.Errors.Add(new ChartLegendEntryMissing(sld, shp, string.Format(AH.A(14179), num, num2), chart.Legend));
				}
			}
		}
		chart = null;
	}

	public static void MissingDataLabels(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		IEnumerator enumerator = ((IEnumerable)shp.Chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
		checked
		{
			try
			{
				IEnumerator enumerator2 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					IMsoSeries msoSeries = (IMsoSeries)enumerator.Current;
					if (!A(msoSeries, shp.Chart))
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
					{
						enumerator2 = ((IEnumerable)msoSeries.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						try
						{
							while (enumerator2.MoveNext())
							{
								IMsoDataLabel msoDataLabel = (IMsoDataLabel)enumerator2.Current;
								if (!A(msoDataLabel))
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
								if (msoDataLabel.ShowValue)
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
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_00e8;
								}
								continue;
								end_IL_00e8:
								break;
							}
						}
						finally
						{
							IDisposable disposable2 = enumerator2 as IDisposable;
							if (disposable2 != null)
							{
								disposable2.Dispose();
							}
						}
					}
					if (num2 > 0)
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
						if (num > 0)
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
							Main.Analysis.Errors.Add(new ChartDataLabelMissing(sld, shp, msoSeries, list));
						}
					}
					list = null;
				}
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

	public static void DataLabelConsistency(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		int num = 0;
		int num2 = 0;
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = ((IEnumerable)shp.Chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
				while (enumerator.MoveNext())
				{
					if (A((IMsoSeries)enumerator.Current, shp.Chart))
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
						num++;
					}
					else
					{
						num2++;
					}
					if (num <= 0 || num2 <= 0)
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
						Main.Analysis.Errors.Add(new ChartDataLabelsInconsistent(sld, shp));
						return;
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

	public static void DataLabelFormat(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
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
				if (!A(msoSeries, shp.Chart))
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				{
					enumerator2 = ((IEnumerable)msoSeries.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
					try
					{
						while (enumerator2.MoveNext())
						{
							IMsoDataLabel msoDataLabel = (IMsoDataLabel)enumerator2.Current;
							if (!A(msoDataLabel))
							{
								continue;
							}
							try
							{
								if (!msoDataLabel.ShowValue)
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
									list.Add(msoDataLabel.NumberFormat);
									break;
								}
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception a = ex;
								Main.A(a, null, shp.Chart);
								ProjectData.ClearProjectError();
							}
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_00fd;
							}
							continue;
							end_IL_00fd:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator2 as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_012c;
				}
				continue;
				end_IL_012c:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		Func<string, D<string>> keySelector;
		if (_Closure_0024__.A == null)
		{
			keySelector = (_Closure_0024__.A = [SpecialName] (string A) => new D<string>(A.ToString()));
		}
		else
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
			keySelector = _Closure_0024__.A;
		}
		IEnumerable<IGrouping<D<string>, string>> enumerable = list.GroupBy(keySelector);
		if (enumerable.Count() > 1)
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
			List<string> list2 = new List<string>();
			List<string> list3 = new List<string>();
			IEnumerator<IGrouping<D<string>, string>> enumerator3 = default(IEnumerator<IGrouping<D<string>, string>>);
			try
			{
				enumerator3 = enumerable.GetEnumerator();
				while (enumerator3.MoveNext())
				{
					IGrouping<D<string>, string> current = enumerator3.Current;
					list3.Add(current.Key.Format + AH.A(14248) + current.Count() + AH.A(14255));
					list2.Add(current.Key.Format);
					current = null;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_0231;
					}
					continue;
					end_IL_0231:
					break;
				}
			}
			finally
			{
				if (enumerator3 != null)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						enumerator3.Dispose();
						break;
					}
				}
			}
			Main.Analysis.Errors.Add(new ChartDataLabelNumberFormats(sld, shp, list3, string.Join(AH.A(14258), list3.ToArray()), list2, shp.Chart.PlotArea));
			list2 = null;
			list3 = null;
		}
		enumerable = null;
	}

	private static bool A(IMsoSeries A, Chart B)
	{
		bool result;
		try
		{
			result = object.Equals(A.HasDataLabels, true);
		}
		catch (NullReferenceException ex)
		{
			ProjectData.SetProjectError(ex);
			NullReferenceException ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception a = ex3;
			Main.A(a, -2147467259, B, new int[8] { 83, 84, 85, 86, 119, 123, 117, 120 });
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static bool A(IMsoDataLabel A)
	{
		bool result;
		try
		{
			_ = A.Text;
			result = true;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}
}
