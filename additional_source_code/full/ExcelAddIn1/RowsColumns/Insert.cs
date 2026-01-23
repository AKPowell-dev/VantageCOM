using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.RowsColumns;

public sealed class Insert
{
	public static void Row()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
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
			Application application = MH.A.Application;
			application.EnableEvents = false;
			try
			{
				application.CutCopyMode = (XlCutCopyMode)0;
				((Range)application.Selection).EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
				Core.LogActivity(VH.A(171787));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			application.EnableEvents = true;
			application = null;
			return;
		}
	}

	public static void Column()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
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
			Application application = MH.A.Application;
			application.EnableEvents = false;
			try
			{
				application.CutCopyMode = (XlCutCopyMode)0;
				((Range)application.Selection).EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight, RuntimeHelpers.GetObjectValue(Missing.Value));
				Core.LogActivity(VH.A(171808));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			application.EnableEvents = true;
			application = null;
			return;
		}
	}
}
