using System;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.View;

public sealed class Zoom
{
	public static void ZoomIn()
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
			try
			{
				Window activeWindow = MH.A.Application.ActiveWindow;
				activeWindow.Zoom = checked(Conversions.ToInteger(activeWindow.Zoom) + KH.A.ZoomInterval);
				_ = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)7, VH.A(175931));
			return;
		}
	}

	public static void ZoomOut()
	{
		if (Licensing.AllowRestrictedMode())
		{
			try
			{
				Window activeWindow = MH.A.Application.ActiveWindow;
				activeWindow.Zoom = checked(Conversions.ToInteger(activeWindow.Zoom) - KH.A.ZoomInterval);
				_ = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)7, VH.A(175931));
		}
	}
}
