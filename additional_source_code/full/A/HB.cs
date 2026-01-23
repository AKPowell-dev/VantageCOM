using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using ExcelAddIn1.Audit.Check.UI;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

internal sealed class HB
{
	[CompilerGenerated]
	private static (string CmdName, List<string> ArgStrs) m_A;

	internal static (string CmdName, List<string> ArgStrs) StartupCmdInfo
	{
		[CompilerGenerated]
		get
		{
			return HB.m_A;
		}
		[CompilerGenerated]
		set
		{
			HB.m_A = value;
		}
	}

	internal static bool A()
	{
		if (!B())
		{
			return C();
		}
		return true;
	}

	private static bool B()
	{
		return true;
	}

	private static bool C()
	{
		try
		{
			Application application = MH.A.Application;
			if (application.Workbooks.Count != 1)
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
						return false;
					}
				}
			}
			return KB.A(application.ActiveWorkbook.FullName);
		}
		finally
		{
			Application application = null;
		}
	}

	internal static void A(bool A)
	{
		bool flag = ExcelAddIn1.Audit.Check.UI.Pane.IsVisible();
		if (flag)
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
			ExcelAddIn1.Audit.Check.UI.Pane.Toggle(blnPressed: false);
		}
		string B = "";
		List<int> C = null;
		if (!HB.A(A, out B, out C))
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (flag)
					{
						ExcelAddIn1.Audit.Check.UI.Pane.Toggle(blnPressed: true);
					}
					return;
				}
			}
		}
		string a = "";
		try
		{
			a = KB.A(B, C);
			Application obj = (Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid(VH.A(7830))));
			obj.Visible = true;
			Application application = obj;
			MH.A.Application.WindowState = XlWindowState.xlMinimized;
			if (application.WindowState == XlWindowState.xlMaximized)
			{
				return;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				application.WindowState = XlWindowState.xlMaximized;
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			KB.A(a);
			Forms.ErrorMessage(string.Format(VH.A(7903), ex2.Message));
			ProjectData.ClearProjectError();
		}
		finally
		{
			Application application = null;
		}
	}

	private static bool A(bool A, out string B, out List<int> C)
	{
		B = "";
		C = null;
		bool result;
		try
		{
			Application application = MH.A.Application;
			Workbooks workbooks = application.Workbooks;
			int count = workbooks.Count;
			if (count < 1)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					result = true;
					break;
				}
			}
			else
			{
				Workbook activeWorkbook = application.ActiveWorkbook;
				B = activeWorkbook.FullName;
				C = (A ? null : LB.A());
				activeWorkbook.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				if (workbooks.Count == count)
				{
					Forms.WarningMessage(VH.A(7978));
					result = false;
				}
				else
				{
					result = true;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(string.Format(VH.A(8118), ex2.Message));
			result = false;
			ProjectData.ClearProjectError();
		}
		finally
		{
			Workbooks workbooks = null;
			Workbook activeWorkbook = null;
			Application application = null;
		}
		return result;
	}
}
