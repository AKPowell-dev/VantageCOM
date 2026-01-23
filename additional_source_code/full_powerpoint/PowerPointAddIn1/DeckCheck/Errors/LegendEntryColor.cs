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

public sealed class LegendEntryColor : BaseColorError
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Action<object, int> A;

		public static Action<object, int> B;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal void A(object A, int B)
		{
			NewLateBinding.LateSet(A, null, AH.A(14136), new object[1] { B }, null, null);
		}

		[SpecialName]
		internal void B(object A, int B)
		{
			NewLateBinding.LateSet(A, null, AH.A(14093), new object[1] { B }, null, null);
		}
	}

	[CompilerGenerated]
	internal sealed class EC
	{
		public bool A;

		public LegendEntryColor A;

		public EC(EC A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal void A(object A, int B)
		{
			if (this.A.m_A)
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
						NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(A, null, AH.A(14028), new object[0], null, null, null), null, AH.A(14041), new object[0], null, null, null), null, AH.A(14050), new object[0], null, null, null), null, AH.A(14069), new object[1] { B }, null, null, OptimisticSet: false, RValueBase: true);
						if (this.A)
						{
							NewLateBinding.LateSet(A, null, AH.A(14136), new object[1] { B }, null, null);
						}
						return;
					}
				}
			}
			NewLateBinding.LateSetComplex(NewLateBinding.LateGet(A, null, AH.A(14076), new object[0], null, null, null), null, AH.A(13587), new object[1] { B }, null, null, OptimisticSet: false, RValueBase: true);
		}
	}

	private new readonly int m_A;

	private new readonly bool m_A;

	private new readonly Chart m_A;

	public LegendEntryColor(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, int intColor, object legendKey, Severity sev, int legendIndex)
		: base(ErrorType.ColorPaletteChartSeries, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		this.m_A = false;
		this.m_A = shp.Chart;
		((BaseError)this).LegendKey = (IMsoLegendKey)legendKey;
		((BaseError)this).LegendEntry = (Microsoft.Office.Core.LegendEntry)NewLateBinding.LateGet(legendKey, null, AH.A(28234), new object[0], null, null, null);
		this.m_A = legendIndex;
		this.m_A = clsCharts.UsesLegendLinesForSeriesClrs(this.m_A);
		((BaseError)this).Title = AH.A(28247);
		((BaseError)this).Subtitle = AH.A(28284);
	}

	public override void FixAction(Color color)
	{
		EC a = default(EC);
		EC CS_0024_003C_003E8__locals8 = new EC(a);
		CS_0024_003C_003E8__locals8.A = this;
		NG.A.Application.StartNewUndoEntry();
		int num = ColorTranslator.ToOle(color);
		CS_0024_003C_003E8__locals8.A = clsCharts.UsesMarkers(this.m_A);
		object objectValue = RuntimeHelpers.GetObjectValue(this.m_A ? ((object)((BaseError)this).LegendKey.Format.Line.ForeColor.RGB) : ((BaseError)this).LegendKey.Interior.Color);
		int num2;
		if (!CS_0024_003C_003E8__locals8.A)
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
			num2 = 0;
		}
		else
		{
			num2 = ((BaseError)this).LegendKey.MarkerBackgroundColor;
		}
		int num3 = num2;
		int num4;
		if (!CS_0024_003C_003E8__locals8.A)
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
			num4 = 0;
		}
		else
		{
			num4 = ((BaseError)this).LegendKey.MarkerForegroundColor;
		}
		int num5 = num4;
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		Dictionary<int, int> dictionary2 = new Dictionary<int, int>();
		Dictionary<int, int> dictionary3 = new Dictionary<int, int>();
		object objectValue2 = RuntimeHelpers.GetObjectValue(A());
		if (objectValue2 != null)
		{
			try
			{
				int num6 = 0;
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = ((IEnumerable)NewLateBinding.LateGet(objectValue2, null, AH.A(23384), new object[0], null, null, null)).GetEnumerator();
					while (enumerator.MoveNext())
					{
						ChartPoint chartPoint = (ChartPoint)enumerator.Current;
						num6 = checked(num6 + 1);
						object obj;
						if (!this.m_A)
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
							obj = chartPoint.Interior.Color;
						}
						else
						{
							obj = chartPoint.Format.Line.ForeColor.RGB;
						}
						object objectValue3 = RuntimeHelpers.GetObjectValue(obj);
						if (!object.Equals(RuntimeHelpers.GetObjectValue(objectValue3), RuntimeHelpers.GetObjectValue(objectValue)))
						{
							dictionary.Add(num6, Conversions.ToInteger(objectValue3));
						}
						if (!CS_0024_003C_003E8__locals8.A)
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
						if (chartPoint.MarkerStyle == XlMarkerStyle.xlMarkerStyleNone)
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
						A(dictionary2, num6, chartPoint.MarkerBackgroundColor, Conversions.ToInteger(objectValue3), Conversions.ToInteger(objectValue), num);
						A(dictionary3, num6, chartPoint.MarkerForegroundColor, Conversions.ToInteger(objectValue3), Conversions.ToInteger(objectValue), num);
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_024e;
						}
						continue;
						end_IL_024e:
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
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		if (this.m_A)
		{
			((BaseError)this).LegendKey.Format.Line.ForeColor.RGB = num;
		}
		else
		{
			((BaseError)this).LegendKey.Interior.Color = num;
		}
		if (CS_0024_003C_003E8__locals8.A)
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
			if (object.Equals(num3, RuntimeHelpers.GetObjectValue(objectValue)))
			{
				((BaseError)this).LegendKey.MarkerBackgroundColor = num;
			}
			if (object.Equals(num5, RuntimeHelpers.GetObjectValue(objectValue)))
			{
				((BaseError)this).LegendKey.MarkerForegroundColor = num;
			}
		}
		A(RuntimeHelpers.GetObjectValue(objectValue2), dictionary, [SpecialName] (object A, int B) =>
		{
			if (CS_0024_003C_003E8__locals8.A.m_A)
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
						NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(A, null, AH.A(14028), new object[0], null, null, null), null, AH.A(14041), new object[0], null, null, null), null, AH.A(14050), new object[0], null, null, null), null, AH.A(14069), new object[1] { B }, null, null, OptimisticSet: false, RValueBase: true);
						if (CS_0024_003C_003E8__locals8.A)
						{
							NewLateBinding.LateSet(A, null, AH.A(14136), new object[1] { B }, null, null);
						}
						return;
					}
				}
			}
			NewLateBinding.LateSetComplex(NewLateBinding.LateGet(A, null, AH.A(14076), new object[0], null, null, null), null, AH.A(13587), new object[1] { B }, null, null, OptimisticSet: false, RValueBase: true);
		});
		object objectValue4 = RuntimeHelpers.GetObjectValue(objectValue2);
		Action<object, int> c;
		if (_Closure_0024__.A == null)
		{
			c = (_Closure_0024__.A = [SpecialName] (object A, int B) =>
			{
				NewLateBinding.LateSet(A, null, AH.A(14136), new object[1] { B }, null, null);
			});
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
			c = _Closure_0024__.A;
		}
		A(objectValue4, dictionary2, c);
		A(RuntimeHelpers.GetObjectValue(objectValue2), dictionary3, [SpecialName] (object A, int B) =>
		{
			NewLateBinding.LateSet(A, null, AH.A(14093), new object[1] { B }, null, null);
		});
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
					switch (6)
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
			object objectValue2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(legendEntries.Cast<object>().ElementAtOrDefault(this.m_A), null, AH.A(13177), new object[0], null, null, null));
			object obj;
			if (!this.m_A)
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
				obj = NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue2, null, AH.A(14076), new object[0], null, null, null), null, AH.A(13587), new object[0], null, null, null);
			}
			else
			{
				obj = NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue2, null, AH.A(14028), new object[0], null, null, null), null, AH.A(14041), new object[0], null, null, null), null, AH.A(14050), new object[0], null, null, null), null, AH.A(14069), new object[0], null, null, null);
			}
			object objectValue3 = RuntimeHelpers.GetObjectValue(obj);
			if (!object.Equals(objB: RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(this.m_A ? ((object)((BaseError)this).LegendKey.Format.Line.ForeColor.RGB) : ((BaseError)this).LegendKey.Interior.Color)), objA: RuntimeHelpers.GetObjectValue(objectValue3)))
			{
				while (true)
				{
					switch (1)
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

	private void A(Dictionary<int, int> A, int B, int C, int D, int E, int F)
	{
		if (new int[2]
		{
			Base.TRANSPARENT,
			D
		}.Contains(C))
		{
			return;
		}
		while (true)
		{
			switch (3)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A.Add(B, (C == E) ? F : C);
			return;
		}
	}

	private void A(object A, Dictionary<int, int> B, Action<object, int> C)
	{
		using Dictionary<int, int>.Enumerator enumerator = B.GetEnumerator();
		while (enumerator.MoveNext())
		{
			KeyValuePair<int, int> current = enumerator.Current;
			object objectValue;
			try
			{
				objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(A, null, AH.A(23384), new object[1] { current.Key }, null, null, null));
				C(RuntimeHelpers.GetObjectValue(objectValue), current.Value);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			objectValue = null;
		}
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
			return;
		}
	}
}
