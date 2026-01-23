using System;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class Underline
{
	[CompilerGenerated]
	private static int m_A;

	internal static int CycleIndex
	{
		[CompilerGenerated]
		get
		{
			return Underline.m_A;
		}
		[CompilerGenerated]
		set
		{
			Underline.m_A = value;
		}
	}

	public static void Cycle()
	{
		checked
		{
			CycleIndex++;
			XlUnderlineStyle a;
			switch (CycleIndex)
			{
			case 1:
				a = XlUnderlineStyle.xlUnderlineStyleSingle;
				break;
			case 2:
				a = XlUnderlineStyle.xlUnderlineStyleSingleAccounting;
				break;
			case 3:
				a = XlUnderlineStyle.xlUnderlineStyleDouble;
				break;
			case 4:
				a = XlUnderlineStyle.xlUnderlineStyleDoubleAccounting;
				break;
			default:
				a = XlUnderlineStyle.xlUnderlineStyleNone;
				CycleIndex = 0;
				break;
			}
			if (CycleIndex == 1)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Base.LogActivity(VH.A(151773));
			}
			A(a);
		}
	}

	public static void DoUnderline(IRibbonControl control)
	{
		A((XlUnderlineStyle)Conversions.ToInteger(control.Tag));
	}

	private static void A(XlUnderlineStyle A)
	{
		Application application = MH.A.Application;
		bool flag = false;
		try
		{
			if (application.Selection is Range)
			{
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
					Range range = JH.A((Range)null);
					if (!Base.IsWorksheetProtected(range.Worksheet))
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
						application.ScreenUpdating = false;
						try
						{
							if (KH.A.UndoFont)
							{
								flag = JH.A(range);
							}
							range.Font.Underline = A;
							if (flag)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									JH.A(range, VH.A(151804));
									break;
								}
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							Base.HandleFormattingException(ex2);
							ProjectData.ClearProjectError();
						}
						application.ScreenUpdating = true;
					}
					range = null;
					break;
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Base.LogException(ex4);
			ProjectData.ClearProjectError();
		}
		application = null;
	}
}
