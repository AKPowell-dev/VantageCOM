using System;
using A;
using MacabacusMacros;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word;

public sealed class clsView
{
	public static void ZoomIn()
	{
		checked
		{
			try
			{
				PC.A.Application.ActiveWindow.View.Zoom.Percentage += 10;
				_ = null;
				clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)7, XC.A(42535));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
	}

	public static void ZoomOut()
	{
		checked
		{
			try
			{
				PC.A.Application.ActiveWindow.View.Zoom.Percentage -= 10;
				_ = null;
				clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)7, XC.A(42550));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
	}
}
