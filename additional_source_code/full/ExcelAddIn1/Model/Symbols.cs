using System;
using A;
using ExcelAddIn1.ExcelApp;
using ExcelAddIn1.Sheets;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Model;

public sealed class Symbols
{
	public static void Insert(int num)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Application application = MH.A.Application;
		if (application.ActiveWindow.SelectedSheets.Count == 1)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (application.ActiveSheet is Worksheet)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					break;
				}
				Worksheet obj = (Worksheet)application.ActiveSheet;
				ExcelAddIn1.Sheets.Protection.Unprotect(obj);
				if (!obj.ProtectContents)
				{
					if (!EditMode.IsEditMode(application))
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
						Range activeCell = default(Range);
						try
						{
							activeCell = application.ActiveCell;
							activeCell.Value2 = Operators.ConcatenateObject(activeCell.Value2, clsText.ConvertUnicodeToChar(num));
							EditMode.EnterEditMode(application);
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							PivotTable pivotTable = null;
							try
							{
								pivotTable = activeCell.PivotTable;
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
							if (pivotTable != null)
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									break;
								}
								pivotTable = null;
								Forms.WarningMessage(VH.A(90774));
							}
							else
							{
								Forms.ErrorMessage(VH.A(90869));
								clsReporting.LogException(ex2);
							}
							ProjectData.ClearProjectError();
						}
						activeCell = null;
						clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(90918));
					}
				}
				else
				{
					Forms.WarningMessage(VH.A(90945));
				}
			}
		}
		else
		{
			Forms.WarningMessage(VH.A(91038));
		}
		application = null;
	}
}
