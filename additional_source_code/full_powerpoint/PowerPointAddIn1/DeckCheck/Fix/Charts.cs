using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Fix;

public sealed class Charts
{
	public static CommandBarControl UndoControl()
	{
		return NG.A.Application.CommandBars[AH.A(47337)].FindControl(RuntimeHelpers.GetObjectValue(Missing.Value), 128, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
	}

	internal static void A(IMsoSeries A, int B, int? C = null, Action<IMsoChartFormat, int, int> D = null, Func<Microsoft.Office.Core.FillFormat, int, int, bool> E = null)
	{
		try
		{
			IMsoChartFormat format = A.Format;
			Microsoft.Office.Core.FillFormat fill = format.Fill;
			int rGB = fill.ForeColor.RGB;
			if (!C.HasValue)
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
				C = rGB;
			}
			if (E == null)
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
				if (!object.Equals(rGB, C.Value))
				{
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
			}
			else if (!E(fill, B, C.Value))
			{
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
			Dictionary<int, int> dictionary = new Dictionary<int, int>();
			if (ImplsPoints(A))
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
				int num = 0;
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = ((IEnumerable)A.Points(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
					while (enumerator.MoveNext())
					{
						object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
						num = checked(num + 1);
						try
						{
							int num2 = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue, null, AH.A(14076), new object[0], null, null, null), null, AH.A(13587), new object[0], null, null, null));
							if (object.Equals(num2, rGB))
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
								dictionary.Add(num, num2);
								break;
							}
						}
						catch (Exception projectError)
						{
							ProjectData.SetProjectError(projectError);
							ProjectData.ClearProjectError();
						}
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_018d;
						}
						continue;
						end_IL_018d:
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
			}
			if (D == null)
			{
				fill.ForeColor.RGB = B;
			}
			else
			{
				D(format, B, C.Value);
			}
			using Dictionary<int, int>.Enumerator enumerator2 = dictionary.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				KeyValuePair<int, int> current = enumerator2.Current;
				object objectValue2 = RuntimeHelpers.GetObjectValue(A.Points(current.Key));
				if (object.Equals(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue2, null, AH.A(14076), new object[0], null, null, null), null, AH.A(13587), new object[0], null, null, null)), current.Value))
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
				NewLateBinding.LateSetComplex(NewLateBinding.LateGet(objectValue2, null, AH.A(14076), new object[0], null, null, null), null, AH.A(13587), new object[1] { current.Value }, null, null, OptimisticSet: false, RValueBase: true);
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
		finally
		{
			Microsoft.Office.Core.FillFormat fill = null;
			IMsoChartFormat format = null;
		}
	}

	public static bool HasRadarAxisLabels(ChartGroup chtGroup)
	{
		bool result;
		try
		{
			result = chtGroup.HasRadarAxisLabels;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool ImplsAndHasErrorBars(IMsoSeries series)
	{
		bool result;
		try
		{
			result = series.HasErrorBars;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool ImplsTrendLines(IMsoSeries series)
	{
		bool result;
		try
		{
			RuntimeHelpers.GetObjectValue(series.Trendlines(RuntimeHelpers.GetObjectValue(Missing.Value)));
			result = true;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool ImplsPoints(IMsoSeries series)
	{
		bool result;
		try
		{
			RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(RuntimeHelpers.GetObjectValue(series.Points(RuntimeHelpers.GetObjectValue(Missing.Value))), null, AH.A(13955), new object[0], null, null, null));
			result = true;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result = false;
			ProjectData.ClearProjectError();
		}
		finally
		{
		}
		return result;
	}

	public static bool ImplsFont(Microsoft.Office.Core.LegendEntry legendEntry)
	{
		bool result;
		try
		{
			_ = legendEntry.Font;
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
