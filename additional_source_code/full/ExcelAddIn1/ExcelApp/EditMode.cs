using System.Windows.Forms;
using A;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.ExcelApp;

public sealed class EditMode
{
	public static bool IsEditMode(Microsoft.Office.Interop.Excel.Application xlApp)
	{
		if (Environment.IsEditMode(xlApp))
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
					Forms.WarningMessage(VH.A(82462));
					return true;
				}
			}
		}
		return false;
	}

	public static void EnterEditMode(Microsoft.Office.Interop.Excel.Application xlApp)
	{
		xlApp.SendKeys(VH.A(49606), false);
		System.Windows.Forms.Application.DoEvents();
		xlApp.SendKeys(VH.A(49606), false);
	}
}
