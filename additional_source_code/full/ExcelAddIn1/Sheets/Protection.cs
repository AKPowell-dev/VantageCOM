using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Sheets;

public sealed class Protection
{
	public static void Unprotect(Worksheet ws)
	{
		if (!ws.ProtectContents)
		{
			return;
		}
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
			try
			{
				ws.Unprotect(RuntimeHelpers.GetObjectValue(Missing.Value));
				return;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(ex2.Message);
				ProjectData.ClearProjectError();
				return;
			}
		}
	}
}
