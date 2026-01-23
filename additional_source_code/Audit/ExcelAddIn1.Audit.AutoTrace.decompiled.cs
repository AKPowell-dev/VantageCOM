using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using A;
using ExcelAddIn1.ExcelApp;
using MacabacusMacros;
using MacabacusMacros.Auth;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit;

public sealed class AutoTrace
{
	private static void A()
	{
		Application application = MH.A.Application;
		if (application.CutCopyMode == (XlCutCopyMode)0)
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
			application.ScreenUpdating = false;
			Arrows.DisplayObjects(application.ActiveWorkbook);
			Range range = default(Range);
			try
			{
				NewLateBinding.LateCall(application.ActiveSheet, null, VH.A(1630), new object[0], null, null, null, IgnoreReturn: true);
				range = (Range)application.Selection;
				if (Operators.ConditionalCompareObjectLessEqual(range.Cells.CountLarge, 20, TextCompare: false))
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						foreach (Range item in range)
						{
							if (!Conversions.ToBoolean(item.HasFormula))
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
							item.ShowPrecedents(RuntimeHelpers.GetObjectValue(Missing.Value));
						}
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
			application.ScreenUpdating = true;
			JH.A((object)range);
		}
		application = null;
	}

	private static void B()
	{
		Application application = MH.A.Application;
		if (application.CutCopyMode == (XlCutCopyMode)0)
		{
			application.ScreenUpdating = false;
			Arrows.DisplayObjects(application.ActiveWorkbook);
			Range range = default(Range);
			try
			{
				NewLateBinding.LateCall(application.ActiveSheet, null, VH.A(1630), new object[0], null, null, null, IgnoreReturn: true);
				range = (Range)application.Selection;
				if (Operators.ConditionalCompareObjectLessEqual(range.Cells.CountLarge, 20, TextCompare: false))
				{
					IEnumerator enumerator = default(IEnumerator);
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
						try
						{
							enumerator = range.GetEnumerator();
							while (enumerator.MoveNext())
							{
								((Range)enumerator.Current).ShowDependents(RuntimeHelpers.GetObjectValue(Missing.Value));
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
			application.ScreenUpdating = true;
			JH.A((object)range);
		}
		application = null;
	}

	public static void TogglePrecedents()
	{
		if (!A())
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					C();
					return;
				}
			}
		}
		Arrows.ClearArrowsOnActiveSheet();
		MySettings settings = K.Settings;
		settings.AutoTracePrecedents = !settings.AutoTracePrecedents;
		if (settings.AutoTracePrecedents)
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
			A();
			settings.AutoTraceDependents = false;
			StatusBar.SetText(VH.A(1653));
			new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).RemoveEventHandler(MH.A.Application, new AppEvents_SheetSelectionChangeEventHandler(AutoTraceDependentsEventHandler));
			new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).AddEventHandler(MH.A.Application, new AppEvents_SheetSelectionChangeEventHandler(AutoTracePrecedentsEventHandler));
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)3, VH.A(1741));
		}
		else
		{
			StatusBar.SetText(VH.A(1782));
			new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).RemoveEventHandler(MH.A.Application, new AppEvents_SheetSelectionChangeEventHandler(AutoTracePrecedentsEventHandler));
		}
		settings = null;
		C();
	}

	public static void ToggleDependents()
	{
		if (!A())
		{
			C();
			return;
		}
		Arrows.ClearArrowsOnActiveSheet();
		MySettings settings = K.Settings;
		settings.AutoTraceDependents = !settings.AutoTraceDependents;
		if (settings.AutoTraceDependents)
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
			B();
			settings.AutoTracePrecedents = false;
			StatusBar.SetText(VH.A(1831));
			new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).RemoveEventHandler(MH.A.Application, new AppEvents_SheetSelectionChangeEventHandler(AutoTracePrecedentsEventHandler));
			new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).AddEventHandler(MH.A.Application, new AppEvents_SheetSelectionChangeEventHandler(AutoTraceDependentsEventHandler));
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)3, VH.A(1878));
		}
		else
		{
			StatusBar.SetText(VH.A(1919));
			new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).RemoveEventHandler(MH.A.Application, new AppEvents_SheetSelectionChangeEventHandler(AutoTraceDependentsEventHandler));
		}
		settings = null;
		C();
	}

	public static void AutoTracePrecedentsEventHandler(object Sh, Range Target)
	{
		A();
	}

	public static void AutoTraceDependentsEventHandler(object Sh, Range Target)
	{
		B();
	}

	private static void C()
	{
		KH.A.InvalidateControl(VH.A(1968));
		KH.A.InvalidateControl(VH.A(1983));
	}

	private static bool A()
	{
		return Access.AllowExcelOperation((PlanType)4, (Restriction)1, false);
	}
}
