using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using A;
using ExcelAddIn1.ExcelApp;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class WrapParentheses
{
	public static void Go()
	{
		if (!Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			if (application.Selection is Range)
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
				Range range = (Range)application.Selection;
				if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
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
					if (Environment.IsEditMode(application))
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
						try
						{
							application.SendKeys(VH.A(156762), false);
							System.Windows.Forms.Application.DoEvents();
							range.Select();
						}
						catch (Exception projectError)
						{
							ProjectData.SetProjectError(projectError);
							ProjectData.ClearProjectError();
						}
					}
					application.ScreenUpdating = false;
					try
					{
						if (Conversions.ToBoolean(Operators.NotObject(range.HasFormula)))
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								throw new Exception();
							}
						}
						bool flag = JH.A(range);
						try
						{
							NewLateBinding.LateSet(range, null, VH.A(1998), new object[1] { Regex.Replace(NewLateBinding.LateGet(range, null, VH.A(1998), new object[0], null, null, null).ToString(), VH.A(156777), VH.A(156792)) + VH.A(39904) }, null, null);
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							range.Formula = Regex.Replace(range.Formula.ToString(), VH.A(156777), VH.A(156792)) + VH.A(39904);
							ProjectData.ClearProjectError();
						}
						if (flag)
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
							JH.A(range, VH.A(155665) + range.get_Address((object)false, (object)false, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
						}
						if (application.EditDirectlyInCell)
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
							EditMode.EnterEditMode(application);
						}
						else
						{
							application.SendKeys(VH.A(49606), true);
						}
						clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(156799));
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
					application.ScreenUpdating = true;
				}
				else
				{
					Forms.WarningMessage(VH.A(156832));
				}
				range = null;
			}
			application = null;
			return;
		}
	}
}
