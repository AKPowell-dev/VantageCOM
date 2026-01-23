using System;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class YAxisFormat
{
	private enum OD
	{
		A,
		B,
		C,
		D,
		E,
		F,
		G,
		H,
		I,
		J,
		K
	}

	public static void ChartAxisMaxFormat(IRibbonControl control)
	{
		Application application = MH.A.Application;
		Axis axis = default(Axis);
		try
		{
			if (application.Selection is Axis)
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
				axis = (Axis)application.Selection;
				goto IL_009d;
			}
			if (!Conversions.ToBoolean(((_Chart)application.ActiveChart).get_HasAxis((object)Microsoft.Office.Core.XlAxisType.xlValue, (object)XlAxisGroup.xlSecondary)))
			{
				axis = (Axis)application.ActiveChart.Axes(Microsoft.Office.Core.XlAxisType.xlValue);
				goto IL_009d;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				Forms.WarningMessage(VH.A(63955));
				break;
			}
			goto end_IL_000f;
			IL_02f0:
			string replacement;
			string text = Regex.Replace(text, VH.A(65110), replacement);
			Axis axis2;
			text = VH.A(65117) + Conversions.ToString(axis2.MaximumScale) + VH.A(43340) + text;
			axis2.TickLabels.NumberFormat = text;
			axis2 = null;
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(65122));
			goto end_IL_000f;
			IL_01c1:
			OD oD;
			switch (oD)
			{
			case OD.A:
				text = VH.A(64027);
				break;
			case OD.B:
				text = VH.A(64110);
				break;
			case OD.C:
				text = VH.A(64197);
				break;
			case OD.D:
				text = VH.A(64284);
				break;
			case OD.F:
				text = VH.A(64375);
				break;
			case OD.H:
				text = VH.A(64466);
				break;
			case OD.I:
				text = VH.A(64543);
				break;
			}
			goto IL_02f0;
			IL_009d:
			axis2 = axis;
			if (axis2.TickLabelPosition == XlTickLabelPosition.xlTickLabelPositionNone)
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
				axis2.TickLabelPosition = XlTickLabelPosition.xlTickLabelPositionNextToAxis;
			}
			axis2.MaximumScaleIsAuto = false;
			MatchCollection matchCollection = Regex.Matches(axis2.TickLabels.NumberFormat, VH.A(64006));
			replacement = ((matchCollection.Count <= 0) ? "" : (VH.A(64021) + application.WorksheetFunction.Rept(VH.A(64024), matchCollection[0].Groups[1].ToString().Length)));
			matchCollection = null;
			oD = (OD)Conversions.ToInteger(control.Tag);
			if (axis.AxisGroup == XlAxisGroup.xlPrimary)
			{
				if (axis.TickLabelPosition != XlTickLabelPosition.xlTickLabelPositionHigh)
				{
					goto IL_01c1;
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
			}
			if (axis.AxisGroup == XlAxisGroup.xlSecondary)
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
				if (axis.TickLabelPosition == XlTickLabelPosition.xlTickLabelPositionLow)
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
					goto IL_01c1;
				}
			}
			switch (oD)
			{
			case OD.A:
				text = VH.A(64616);
				break;
			case OD.B:
				text = VH.A(64695);
				break;
			case OD.D:
				text = VH.A(64778);
				break;
			case OD.F:
				text = VH.A(64865);
				break;
			case OD.H:
				text = VH.A(64952);
				break;
			case OD.I:
				text = VH.A(65033);
				break;
			}
			goto IL_02f0;
			end_IL_000f:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application = null;
		JH.A((object)axis);
	}

	public static void ChartAxisFormat(IRibbonControl control)
	{
		Application application = MH.A.Application;
		Axis axis = null;
		string text = string.Empty;
		try
		{
			if (application.Selection is Axis)
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
				axis = (Axis)application.Selection;
				goto IL_009f;
			}
			if (!Conversions.ToBoolean(((_Chart)application.ActiveChart).get_HasAxis((object)Microsoft.Office.Core.XlAxisType.xlValue, (object)XlAxisGroup.xlSecondary)))
			{
				axis = (Axis)application.ActiveChart.Axes(Microsoft.Office.Core.XlAxisType.xlValue);
				goto IL_009f;
			}
			Forms.WarningMessage(VH.A(63955));
			goto end_IL_0017;
			IL_009f:
			Axis axis2 = axis;
			if (axis2.TickLabelPosition == XlTickLabelPosition.xlTickLabelPositionNone)
			{
				axis2.TickLabelPosition = XlTickLabelPosition.xlTickLabelPositionNextToAxis;
			}
			OD oD = (OD)Conversions.ToInteger(control.Tag);
			if (oD != OD.J)
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
				if (oD != OD.K)
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
				}
				else
				{
					text = VH.A(65220);
				}
			}
			else
			{
				text = VH.A(65157);
			}
			if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
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
				axis2.TickLabels.NumberFormat = text;
			}
			axis2 = null;
			end_IL_0017:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application = null;
		JH.A((object)axis);
	}

	public static void RescaleAxis()
	{
		try
		{
			object instance = MH.A.Application.ActiveChart.Axes(Microsoft.Office.Core.XlAxisType.xlValue);
			NewLateBinding.LateSetComplex(instance, null, VH.A(65275), new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
			object instance2 = NewLateBinding.LateGet(instance, null, VH.A(60400), new object[0], null, null, null);
			string memberName = VH.A(57240);
			object[] array = new object[1];
			Type typeFromHandle = typeof(Regex);
			string memberName2 = VH.A(65312);
			object instance3;
			object[] obj = new object[3]
			{
				NewLateBinding.LateGet(instance3 = NewLateBinding.LateGet(instance, null, VH.A(60400), new object[0], null, null, null), null, VH.A(57240), new object[0], null, null, null),
				VH.A(65327),
				Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(65117), NewLateBinding.LateGet(instance, null, VH.A(65354), new object[0], null, null, null)), VH.A(43340))
			};
			object[] array2 = obj;
			bool[] obj2 = new bool[3] { true, false, false };
			bool[] array3 = obj2;
			object obj3 = NewLateBinding.LateGet(null, typeFromHandle, memberName2, obj, null, null, obj2);
			if (array3[0])
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				NewLateBinding.LateSetComplex(instance3, null, VH.A(57240), new object[1] { array2[0] }, null, null, OptimisticSet: true, RValueBase: true);
			}
			array[0] = obj3;
			NewLateBinding.LateSetComplex(instance2, null, memberName, array, null, null, OptimisticSet: false, RValueBase: true);
			NewLateBinding.LateSetComplex(instance, null, VH.A(65275), new object[1] { false }, null, null, OptimisticSet: false, RValueBase: true);
			instance = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
