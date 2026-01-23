using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using A;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.View;

public sealed class NavAid
{
	private static readonly string m_A = VH.A(174785);

	private static readonly string m_B = VH.A(174812);

	private static readonly string m_C = VH.A(174839);

	private static readonly string D = VH.A(174866);

	public static readonly int PATTERN_COLOR_GREEN = ColorTranslator.ToOle(Color.FromArgb(210, 240, 224));

	private static readonly int m_A = ColorTranslator.ToOle(Color.Transparent);

	private static bool m_A = false;

	private static string E = string.Empty;

	private static List<Worksheet> m_A = null;

	private static long m_A = 0L;

	private static long m_B = 0L;

	public static bool Enabled
	{
		get
		{
			return NavAid.m_A;
		}
		set
		{
			NavAid.m_A = value;
		}
	}

	public static string ConditionalFormula
	{
		get
		{
			return E;
		}
		set
		{
			E = value;
		}
	}

	private static List<Worksheet> A
	{
		get
		{
			return NavAid.m_A;
		}
		set
		{
			NavAid.m_A = value;
		}
	}

	private static long A
	{
		get
		{
			return NavAid.m_A;
		}
		set
		{
			NavAid.m_A = value;
		}
	}

	private static long B
	{
		get
		{
			return NavAid.m_B;
		}
		set
		{
			NavAid.m_B = value;
		}
	}

	public static void Toggle(bool blnChecked)
	{
		if (blnChecked)
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
					A();
					return;
				}
			}
		}
		Disable();
	}

	private static void A()
	{
		if (!Licensing.AllowAdvancedViewOperation())
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					RefreshRibbon();
					return;
				}
			}
		}
		Enabled = true;
		NavAid.A = new List<Worksheet>();
		Application application = MH.A.Application;
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).AddEventHandler(application, new AppEvents_SheetSelectionChangeEventHandler(A));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(125628)).AddEventHandler(application, new AppEvents_SheetActivateEventHandler(A));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(47767)).AddEventHandler(application, new AppEvents_WorkbookActivateEventHandler(A));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(125655)).AddEventHandler(application, new AppEvents_WorkbookBeforeSaveEventHandler(A));
		if (application.ActiveSheet is Worksheet)
		{
			NavAid.A.Add((Worksheet)application.ActiveSheet);
		}
		application.ScreenUpdating = false;
		try
		{
			Range obj = (Range)application.Selection;
			A(obj);
			B(obj.Worksheet);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		finally
		{
		}
		application.ScreenUpdating = true;
		application = null;
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)7, VH.A(174600));
	}

	public static void Disable()
	{
		Application application = MH.A.Application;
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).RemoveEventHandler(application, new AppEvents_SheetSelectionChangeEventHandler(A));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(125628)).RemoveEventHandler(application, new AppEvents_SheetActivateEventHandler(A));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(47767)).RemoveEventHandler(application, new AppEvents_WorkbookActivateEventHandler(A));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(125655)).RemoveEventHandler(application, new AppEvents_WorkbookBeforeSaveEventHandler(A));
		application.ScreenUpdating = false;
		using (List<Worksheet>.Enumerator enumerator = NavAid.A.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				Worksheet current = enumerator.Current;
				try
				{
					C(current);
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
		try
		{
			B();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		application.ScreenUpdating = true;
		application = null;
		Enabled = false;
		JH.A((object)NavAid.A);
	}

	public static void RefreshRibbon()
	{
		KH.A.InvalidateControl(VH.A(174613));
	}

	private static void A(object A, Range B)
	{
		Application application = MH.A.Application;
		application.ScreenUpdating = false;
		try
		{
			NavAid.A(B);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application.ScreenUpdating = true;
		application = null;
	}

	private static void A(object A)
	{
		if (!(A is Worksheet))
		{
			return;
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
			NavAid.A((Worksheet)A);
			return;
		}
	}

	private static void A(Microsoft.Office.Interop.Excel.Workbook A)
	{
		if (A.ActiveSheet is Worksheet)
		{
			NavAid.A((Worksheet)A.ActiveSheet);
		}
	}

	private static void A(Worksheet A)
	{
		NavAid.A((Range)A.Application.Selection);
		if (NavAid.A.Contains(A))
		{
			return;
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			NavAid.A.Add(A);
			B(A);
			return;
		}
	}

	private static void A(Microsoft.Office.Interop.Excel.Workbook A, bool B, ref bool C)
	{
		Disable();
		RefreshRibbon();
	}

	private static void A(Range A)
	{
		string text = "";
		string text2 = "";
		string text3 = "";
		string text4 = "";
		if (NavAid.A == 0L)
		{
			NavAid.A = Conversions.ToLong(A.Worksheet.Rows.CountLarge);
		}
		if (NavAid.B == 0L)
		{
			NavAid.B = Conversions.ToLong(A.Worksheet.Columns.CountLarge);
		}
		if (Operators.ConditionalCompareObjectLess(A.Rows.CountLarge, NavAid.A, TextCompare: false))
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
			if (Operators.ConditionalCompareObjectLess(A.Columns.CountLarge, NavAid.B, TextCompare: false))
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
				{
					IEnumerator enumerator = A.Areas.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							Range range = (Range)enumerator.Current;
							text = text + Conversions.ToString(range.Row) + VH.A(2378);
							text2 = Conversions.ToString(Operators.ConcatenateObject(text2, Operators.ConcatenateObject(Operators.SubtractObject(Operators.AddObject(range.Row, range.Rows.CountLarge), 1), VH.A(2378))));
							text3 = text3 + Conversions.ToString(range.Column) + VH.A(2378);
							text4 = Conversions.ToString(Operators.ConcatenateObject(text4, Operators.ConcatenateObject(Operators.SubtractObject(Operators.AddObject(range.Column, range.Columns.CountLarge), 1), VH.A(2378))));
							range = null;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_01e2;
							}
							continue;
							end_IL_01e2:
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
				goto IL_0252;
			}
		}
		text += VH.A(174638);
		text2 += VH.A(174638);
		text3 += VH.A(174638);
		text4 += VH.A(174638);
		goto IL_0252;
		IL_0252:
		Names names = ((Microsoft.Office.Interop.Excel.Workbook)A.Worksheet.Parent).Names;
		checked
		{
			names.Add(NavAid.m_A, VH.A(174643) + Strings.Left(text, text.Length - 1) + VH.A(19802), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Visible = false;
			_ = null;
			names.Add(NavAid.m_B, VH.A(174643) + Strings.Left(text2, text2.Length - 1) + VH.A(19802), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Visible = false;
			_ = null;
			names.Add(NavAid.m_C, VH.A(174643) + Strings.Left(text3, text3.Length - 1) + VH.A(19802), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Visible = false;
			_ = null;
			names.Add(D, VH.A(174643) + Strings.Left(text4, text4.Length - 1) + VH.A(19802), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Visible = false;
			_ = null;
			_ = null;
		}
	}

	private static void B()
	{
		IEnumerator enumerator = MH.A.Application.Workbooks.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)enumerator.Current;
				try
				{
					Names names = workbook.Names;
					names.Item(NavAid.m_A, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Delete();
					names.Item(NavAid.m_B, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Delete();
					names.Item(NavAid.m_C, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Delete();
					names.Item(D, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Delete();
					_ = null;
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
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	private static void B(Worksheet A)
	{
		FormatConditions formatConditions = A.Rows.FormatConditions;
		formatConditions.Add(XlFormatConditionType.xlExpression, RuntimeHelpers.GetObjectValue(Missing.Value), NavAid.A(), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		object instance = NewLateBinding.LateGet(formatConditions.Item(formatConditions.Count), null, VH.A(36170), new object[0], null, null, null);
		NewLateBinding.LateSetComplex(instance, null, VH.A(36187), new object[1] { PATTERN_COLOR_GREEN }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(instance, null, VH.A(36212), new object[1] { XlPattern.xlPatternGray50 }, null, null, OptimisticSet: false, RValueBase: true);
		instance = null;
		_ = null;
	}

	private static void C(Worksheet A)
	{
		FormatConditions formatConditions = A.Rows.FormatConditions;
		for (int i = formatConditions.Count; i >= 1; i = checked(i + -1))
		{
			try
			{
				FormatCondition formatCondition = (FormatCondition)formatConditions.Item(i);
				if (formatCondition.Type == 2)
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
					if (Operators.CompareString(formatCondition.Formula1, NavAid.A(), TextCompare: false) == 0)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								formatCondition.Delete();
								goto end_IL_008a;
							}
						}
					}
				}
				formatCondition = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			continue;
			end_IL_008a:
			break;
		}
		formatConditions = null;
	}

	private static string A()
	{
		if (Operators.CompareString(ConditionalFormula, string.Empty, TextCompare: false) == 0)
		{
			string text;
			string text2;
			string text3;
			string text4;
			if (MH.A.Application.LanguageSettings.get_LanguageID(MsoAppLanguageID.msoLanguageIDUI) == 1033)
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
				text = VH.A(174648);
				text2 = VH.A(47446);
				text3 = VH.A(47410);
				text4 = VH.A(174653);
			}
			else
			{
				Name name = MH.A.Application.ActiveWorkbook.Names.Add(VH.A(94040), VH.A(174674), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				name.RefersTo = VH.A(174679);
				text = A(name.RefersToLocal.ToString());
				name.RefersTo = VH.A(148364);
				text2 = A(name.RefersToLocal.ToString());
				name.RefersTo = VH.A(148381);
				text3 = A(name.RefersToLocal.ToString());
				name.RefersTo = VH.A(174692);
				text4 = A(name.RefersToLocal.ToString());
				_ = null;
				name.Delete();
			}
			ConditionalFormula = VH.A(48936) + text + VH.A(39848) + text4 + VH.A(174725) + text2 + VH.A(174730) + NavAid.m_A + VH.A(174743) + text2 + VH.A(174750) + NavAid.m_B + VH.A(174763) + text4 + VH.A(174725) + text3 + VH.A(174730) + NavAid.m_C + VH.A(174743) + text3 + VH.A(174750) + D + VH.A(174774);
		}
		return ConditionalFormula;
	}

	private static string A(string A)
	{
		return Strings.Mid(A, 2, checked(A.IndexOf(VH.A(39848)) - 1));
	}
}
