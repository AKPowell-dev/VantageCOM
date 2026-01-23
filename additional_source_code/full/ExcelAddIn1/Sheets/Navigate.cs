using System;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Sheets;

public sealed class Navigate
{
	internal static void A()
	{
		try
		{
			NewLateBinding.LateCall(MH.A.Application.ActiveWorkbook.Sheets[1], null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(102151));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	internal static void B()
	{
		try
		{
			Microsoft.Office.Interop.Excel.Workbook activeWorkbook = MH.A.Application.ActiveWorkbook;
			NewLateBinding.LateCall(activeWorkbook.Sheets[activeWorkbook.Sheets.Count], null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
			activeWorkbook = null;
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(102151));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	internal static void C()
	{
		Application application = MH.A.Application;
		try
		{
			int num = Conversions.ToInteger(Operators.AddObject(NewLateBinding.LateGet(application.ActiveSheet, null, VH.A(48135), new object[0], null, null, null), 1));
			int count = application.Sheets.Count;
			for (int i = num; i <= count; i = checked(i + 1))
			{
				if (!Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(application.Sheets[i], null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetVisible, TextCompare: false))
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					NewLateBinding.LateCall(application.Sheets[i], null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
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
					goto end_IL_00e8;
				}
				continue;
				end_IL_00e8:
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application = null;
		A();
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(102151));
	}

	internal static void D()
	{
		Application application = MH.A.Application;
		try
		{
			for (int i = Conversions.ToInteger(Operators.SubtractObject(NewLateBinding.LateGet(application.ActiveSheet, null, VH.A(48135), new object[0], null, null, null), 1)); i >= 1; i = checked(i + -1))
			{
				if (!Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(application.Sheets[i], null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetVisible, TextCompare: false))
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					NewLateBinding.LateCall(application.Sheets[i], null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
					return;
				}
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_00da;
				}
				continue;
				end_IL_00da:
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application = null;
		B();
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(102151));
	}

	internal static void E()
	{
		if (!Licensing.AllowRestrictedMode())
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
			new wpfSheetActivate().ShowDialog();
			_ = null;
			return;
		}
	}
}
