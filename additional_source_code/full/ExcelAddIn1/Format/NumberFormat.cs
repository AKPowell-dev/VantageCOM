using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.ExcelApp;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class NumberFormat
{
	public static void CycleNumber()
	{
		A(KH.A.CycleNumber);
	}

	public static void CycleCurrency()
	{
		A(KH.A.CycleCurrency);
	}

	public static void CyclePercent()
	{
		A(KH.A.CyclePercent);
	}

	public static void CycleMultiple()
	{
		A(KH.A.CycleMultiple);
	}

	public static void CycleDate()
	{
		A(KH.A.CycleDate);
	}

	public static void CycleBinary()
	{
		A(KH.A.CycleBinary);
	}

	public static void CycleRatio()
	{
		A(KH.A.CycleRatio);
	}

	public static void CycleForeign()
	{
		Forms.WarningMessage(VH.A(150630));
	}

	private static void A(NumberFormatCycle A)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		checked
		{
			while (true)
			{
				switch (5)
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
					NumberFormatCycle numberFormatCycle = A;
					int count = numberFormatCycle.Items.Count;
					if (count > 0)
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
						NumberFormat.A(A, numberFormatCycle.Index);
						if (numberFormatCycle.Index == 0)
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
							Base.LogActivity(numberFormatCycle.Activity);
						}
						if (numberFormatCycle.Index < count - 1)
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
							numberFormatCycle.Index++;
						}
						else
						{
							numberFormatCycle.Index = 0;
						}
					}
					numberFormatCycle = null;
					return;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
					return;
				}
			}
		}
	}

	public static void DoGeneral(IRibbonControl control)
	{
		A(KH.A.CycleNumber, Conversions.ToInteger(control.Tag));
	}

	public static void DoCurrency(IRibbonControl control)
	{
		A(KH.A.CycleCurrency, Conversions.ToInteger(control.Tag));
	}

	public static void DoPercent(IRibbonControl control)
	{
		A(KH.A.CyclePercent, Conversions.ToInteger(control.Tag));
	}

	public static void DoMultiple(IRibbonControl control)
	{
		A(KH.A.CycleMultiple, Conversions.ToInteger(control.Tag));
	}

	public static void DoDate(IRibbonControl control)
	{
		A(KH.A.CycleDate, Conversions.ToInteger(control.Tag));
	}

	public static void DoBinary(IRibbonControl control)
	{
		A(KH.A.CycleBinary, Conversions.ToInteger(control.Tag));
	}

	public static void DoRatio(IRibbonControl control)
	{
		A(KH.A.CycleRatio, Conversions.ToInteger(control.Tag));
	}

	private static void A(NumberFormatCycle A, int B)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		try
		{
			if (application.Selection is Range)
			{
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					application.ScreenUpdating = false;
					try
					{
						bool num = JH.A(range);
						string text = (string)(range.NumberFormat = A.Items[B].Format);
						if (KH.A.AutoAlignRightNumbers)
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
							range.HorizontalAlignment = XlHAlign.xlHAlignRight;
						}
						if (KH.A.AutoItalicizePercentages)
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
							if (!text.Contains(VH.A(75697)) && !text.Contains(VH.A(150932)))
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
								if (!text.Contains(VH.A(150937)))
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
									if (!text.Contains(VH.A(150944)))
									{
										goto IL_0143;
									}
									while (true)
									{
										switch (1)
										{
										case 0:
											continue;
										}
										break;
									}
								}
							}
							range.Font.Italic = true;
						}
						goto IL_0143;
						IL_0143:
						if (num)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								JH.A(range, VH.A(148068));
								break;
							}
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						HandleException(ex2);
						ProjectData.ClearProjectError();
					}
					application.ScreenUpdating = true;
				}
				range = null;
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

	public static void HandleException(Exception ex)
	{
		if (ex is COMException)
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
			if (ex.Message.ToLower().Contains(VH.A(148573)))
			{
				MessageBox.Show(ex.Message, VH.A(43304), MessageBoxButtons.OK, MessageBoxIcon.Hand);
				return;
			}
		}
		if (!ex.Message.ToLower().Contains(VH.A(150951)) || EditMode.IsEditMode(MH.A.Application))
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			Forms.ErrorMessage(VH.A(151068));
			return;
		}
	}
}
