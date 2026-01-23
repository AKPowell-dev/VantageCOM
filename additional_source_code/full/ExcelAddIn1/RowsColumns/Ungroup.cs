using System;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.RowsColumns;

public sealed class Ungroup
{
	public static void Rows()
	{
		if (Licensing.AllowRestrictedMode())
		{
			Application application = MH.A.Application;
			application.ScreenUpdating = false;
			try
			{
				Unhide.Rows(blnLogActivity: false);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			try
			{
				((Range)application.Selection).Rows.Ungroup();
				Core.LogActivity(VH.A(171954));
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			application.ScreenUpdating = true;
			application = null;
		}
	}

	public static void Columns()
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
			Application application = MH.A.Application;
			application.ScreenUpdating = false;
			try
			{
				Unhide.Columns(blnLogActivity: false);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			try
			{
				((Range)application.Selection).Columns.Ungroup();
				Core.LogActivity(VH.A(171977));
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			application.ScreenUpdating = true;
			application = null;
			return;
		}
	}
}
