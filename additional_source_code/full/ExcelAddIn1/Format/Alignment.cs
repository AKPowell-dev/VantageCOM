using System;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Formulas;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class Alignment
{
	[CompilerGenerated]
	private static int m_A;

	[CompilerGenerated]
	private static int B;

	internal static int HorizontalCycleIndex
	{
		[CompilerGenerated]
		get
		{
			return Alignment.m_A;
		}
		[CompilerGenerated]
		set
		{
			Alignment.m_A = value;
		}
	}

	internal static int VerticalCycleIndex
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public static void CycleCenter()
	{
		Forms.WarningMessage(VH.A(147515));
	}

	public static void CycleHorizontal()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		checked
		{
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
				HorizontalCycleIndex++;
				XlHAlign a;
				switch (HorizontalCycleIndex)
				{
				case 1:
					a = XlHAlign.xlHAlignCenter;
					break;
				case 2:
					a = XlHAlign.xlHAlignCenterAcrossSelection;
					break;
				case 3:
					a = XlHAlign.xlHAlignRight;
					break;
				case 4:
					a = XlHAlign.xlHAlignLeft;
					break;
				default:
					a = XlHAlign.xlHAlignGeneral;
					HorizontalCycleIndex = 0;
					break;
				}
				if (HorizontalCycleIndex == 1)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						break;
					}
					Base.LogActivity(VH.A(147785));
				}
				A(a);
				return;
			}
		}
	}

	public static void CycleVertical()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		checked
		{
			VerticalCycleIndex++;
			XlVAlign a;
			switch (VerticalCycleIndex)
			{
			case 1:
				a = XlVAlign.xlVAlignCenter;
				break;
			case 2:
				a = XlVAlign.xlVAlignTop;
				break;
			default:
				a = XlVAlign.xlVAlignBottom;
				VerticalCycleIndex = 0;
				break;
			}
			if (VerticalCycleIndex == 1)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Base.LogActivity(VH.A(147818));
			}
			A(a);
		}
	}

	public static void DoAlignHorizontal(IRibbonControl control)
	{
		A((XlHAlign)Conversions.ToInteger(control.Tag));
	}

	public static void DoAlignVertical(IRibbonControl control)
	{
		A((XlVAlign)Conversions.ToInteger(control.Tag));
	}

	private static void A(XlHAlign A)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		bool flag = false;
		bool flag2 = false;
		try
		{
			if (application.Selection is Range)
			{
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
					Range range = JH.A((Range)null);
					if (!Base.IsWorksheetProtected(range.Worksheet))
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
						application.ScreenUpdating = false;
						try
						{
							if (A == XlHAlign.xlHAlignCenterAcrossSelection)
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
								if (Helpers.ContainsMergedCells(range))
								{
									while (true)
									{
										switch (5)
										{
										case 0:
											continue;
										}
										break;
									}
									if (System.Windows.Forms.MessageBox.Show(VH.A(147847), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
									{
										if (!Workbooks.IsShared((Microsoft.Office.Interop.Excel.Workbook)range.Worksheet.Parent, true, (System.Windows.Window)null))
										{
											range.UnMerge();
										}
										else
										{
											flag2 = true;
										}
									}
									else
									{
										flag2 = true;
									}
								}
							}
							if (!flag2)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									if (KH.A.UndoAlignment)
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
										flag = JH.A(range);
									}
									range.HorizontalAlignment = A;
									if (!flag)
									{
										break;
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											continue;
										}
										JH.A(range, VH.A(148035));
										break;
									}
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

	private static void A(XlVAlign A)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
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
						application.ScreenUpdating = false;
						try
						{
							if (KH.A.UndoAlignment)
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
								flag = JH.A(range);
							}
							range.VerticalAlignment = A;
							if (flag)
							{
								JH.A(range, VH.A(148068));
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
