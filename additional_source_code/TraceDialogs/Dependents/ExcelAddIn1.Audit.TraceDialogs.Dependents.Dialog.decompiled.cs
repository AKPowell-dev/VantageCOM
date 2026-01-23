using System;
using System.Linq;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.TraceDialogs.Dependents;

public sealed class Dialog
{
	public static void Show()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		if (Base.CanTrace(application))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Base.RecordLastAuditedCell(application);
			if (System.Windows.Forms.Application.OpenForms.OfType<frmDependentsHost>().Any())
			{
				System.Windows.Forms.Application.OpenForms.OfType<frmDependentsHost>().ElementAt(0).Close();
			}
			Base.DisableNavAid();
			Base.InitializeRegex();
			frmDependentsHost frmDependentsHost = new frmDependentsHost();
			try
			{
				frmDependentsHost.Show();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			frmDependentsHost = null;
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)3, VH.A(41603));
		}
		application = null;
	}

	public static void ClearArrowsOnActiveSheet(Microsoft.Office.Interop.Excel.Application xlApp)
	{
		try
		{
			NewLateBinding.LateCall(xlApp.ActiveSheet, null, VH.A(1630), new object[0], null, null, null, IgnoreReturn: true);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
