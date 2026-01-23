using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Colors;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class LegendMarkerColor : BaseColorError
{
	private new readonly int m_A;

	private new readonly bool? m_A;

	private new readonly Chart m_A;

	public LegendMarkerColor(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, int intColor, object legendKey, Severity sev, int legendIndex, bool? isForeClr)
		: base(ErrorType.ColorPaletteChartSeries, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		this.m_A = shp.Chart;
		((BaseError)this).LegendKey = (IMsoLegendKey)legendKey;
		this.m_A = legendIndex;
		this.m_A = isForeClr;
		if (!isForeClr.HasValue)
		{
			((BaseError)this).Title = AH.A(23397);
			((BaseError)this).Subtitle = AH.A(23436);
		}
		else if (isForeClr.Value)
		{
			((BaseError)this).Title = AH.A(23563);
			((BaseError)this).Subtitle = AH.A(23604);
		}
		else
		{
			((BaseError)this).Title = AH.A(23734);
			((BaseError)this).Subtitle = AH.A(23771);
		}
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		int a = ColorTranslator.ToOle(color);
		if (this.m_A.HasValue)
		{
			if (!this.m_A.Value)
			{
				goto IL_0057;
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
		}
		A(a, B: true);
		goto IL_0057;
		IL_0057:
		if (!this.m_A.HasValue || !this.m_A.Value)
		{
			A(a, B: false);
		}
	}

	private void A(int A, bool B)
	{
		int num;
		if (!B)
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
			num = ((BaseError)this).LegendKey.MarkerBackgroundColor;
		}
		else
		{
			num = ((BaseError)this).LegendKey.MarkerForegroundColor;
		}
		int num2 = num;
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		object objectValue = RuntimeHelpers.GetObjectValue(this.A());
		if (objectValue != null)
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
			try
			{
				int num3 = 0;
				IEnumerator enumerator = ((IEnumerable)NewLateBinding.LateGet(objectValue, null, AH.A(23384), new object[0], null, null, null)).GetEnumerator();
				try
				{
					while (enumerator.MoveNext())
					{
						ChartPoint chartPoint = (ChartPoint)enumerator.Current;
						num3 = checked(num3 + 1);
						if (chartPoint.MarkerStyle == XlMarkerStyle.xlMarkerStyleNone)
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
						int num4;
						if (!B)
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
							num4 = chartPoint.MarkerBackgroundColor;
						}
						else
						{
							num4 = chartPoint.MarkerForegroundColor;
						}
						int num5 = num4;
						if (num5 == Base.TRANSPARENT)
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
						if (object.Equals(num5, num2))
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
						dictionary.Add(num3, num5);
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_0123;
						}
						continue;
						end_IL_0123:
						break;
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
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		if (B)
		{
			((BaseError)this).LegendKey.MarkerForegroundColor = A;
		}
		else
		{
			((BaseError)this).LegendKey.MarkerBackgroundColor = A;
		}
		using Dictionary<int, int>.Enumerator enumerator2 = dictionary.GetEnumerator();
		while (enumerator2.MoveNext())
		{
			KeyValuePair<int, int> current = enumerator2.Current;
			object objectValue2;
			try
			{
				objectValue2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, null, AH.A(23384), new object[1] { current.Key }, null, null, null));
				if (B)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						NewLateBinding.LateSet(objectValue2, null, AH.A(14093), new object[1] { current.Value }, null, null);
						break;
					}
				}
				else
				{
					NewLateBinding.LateSet(objectValue2, null, AH.A(14136), new object[1] { current.Value }, null, null);
				}
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				ProjectData.ClearProjectError();
			}
			objectValue2 = null;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	private object A()
	{
		try
		{
			LegendEntries legendEntries = (LegendEntries)this.m_A.Legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value));
			object objectValue = RuntimeHelpers.GetObjectValue(this.m_A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)));
			if (Operators.ConditionalCompareObjectNotEqual(NewLateBinding.LateGet(objectValue, null, AH.A(13955), new object[0], null, null, null), legendEntries.Count, TextCompare: false))
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return null;
					}
				}
			}
			if (this.m_A >= legendEntries.Count)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						return null;
					}
				}
			}
			if (!object.Equals(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(legendEntries.Cast<object>().ElementAtOrDefault(this.m_A), null, AH.A(13177), new object[0], null, null, null)), null, AH.A(14136), new object[0], null, null, null)), ((BaseError)this).LegendKey.MarkerBackgroundColor))
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						return null;
					}
				}
			}
			return NewLateBinding.LateIndexGet(objectValue, new object[1] { checked(this.m_A + 1) }, null);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		finally
		{
			object objectValue = null;
			LegendEntries legendEntries = null;
		}
		return null;
	}
}
