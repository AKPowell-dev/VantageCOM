using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.ExcelApp;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Sheets;

public sealed class Visibility
{
	internal static bool A(object A)
	{
		if (Operators.ConditionalCompareObjectNotEqual(NewLateBinding.LateGet(A, null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetVisible, TextCompare: false) && Forms.OkCancelMessage2(VH.A(102182)) == DialogResult.OK)
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
			NewLateBinding.LateSet(A, null, VH.A(41367), new object[1] { XlSheetVisibility.xlSheetVisible }, null, null);
		}
		return Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(A, null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetVisible, TextCompare: false);
	}

	internal static void A()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		try
		{
			Microsoft.Office.Interop.Excel.Application application2 = application;
			if (application2.Workbooks.Count > 0)
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
				if (application2.ActiveWorkbook.Sheets.Count > 1)
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
					if (!EditMode.IsEditMode(application))
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
						IEnumerator enumerator = default(IEnumerator);
						try
						{
							enumerator = application2.ActiveWindow.SelectedSheets.GetEnumerator();
							while (enumerator.MoveNext())
							{
								object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
								if (objectValue is Worksheet)
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
									((Worksheet)objectValue).Visible = XlSheetVisibility.xlSheetVeryHidden;
								}
								else
								{
									if (!(objectValue is Chart))
									{
										continue;
									}
									while (true)
									{
										switch (6)
										{
										case 0:
											continue;
										}
										break;
									}
									((Chart)objectValue).Visible = XlSheetVisibility.xlSheetVeryHidden;
								}
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_00ef;
								}
								continue;
								end_IL_00ef:
								break;
							}
						}
						finally
						{
							if (enumerator is IDisposable)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									(enumerator as IDisposable).Dispose();
									break;
								}
							}
						}
						clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(102279));
					}
				}
				else
				{
					Forms.WarningMessage(VH.A(102306));
				}
			}
			application2 = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		application = null;
	}

	internal static void B()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
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
			int C = 0;
			A(XlSheetVisibility.xlSheetHidden, XlSheetVisibility.xlSheetVeryHidden, ref C);
			Forms.InfoMessage(VH.A(102393) + C + VH.A(102408));
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(102279));
			return;
		}
	}

	internal static void C()
	{
		if (!Licensing.AllowRestrictedMode())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			int C = 0;
			A(XlSheetVisibility.xlSheetVeryHidden, XlSheetVisibility.xlSheetVisible, ref C);
			Forms.InfoMessage(VH.A(102443) + C + VH.A(102458));
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(102279));
			return;
		}
	}

	private static void A(XlSheetVisibility A, XlSheetVisibility B, ref int C)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		checked
		{
			if (application.Workbooks.Count > 0)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				try
				{
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = application.ActiveWorkbook.Sheets.GetEnumerator();
						while (enumerator.MoveNext())
						{
							object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
							if (objectValue is Worksheet)
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
								Worksheet worksheet = (Worksheet)objectValue;
								if (worksheet.Visible == A && Operators.CompareString(worksheet.Name, clsDiscuss.HIDDEN_SHEET_NAME, TextCompare: false) != 0)
								{
									worksheet.Visible = B;
									C++;
								}
								worksheet = null;
							}
							else
							{
								if (!(objectValue is Chart))
								{
									continue;
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									break;
								}
								Chart chart = (Chart)objectValue;
								if (chart.Visible == A)
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
									chart.Visible = B;
									C++;
								}
								chart = null;
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_00f7;
							}
							continue;
							end_IL_00f7:
							break;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								(enumerator as IDisposable).Dispose();
								break;
							}
						}
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					clsReporting.LogException(ex2);
					ProjectData.ClearProjectError();
				}
			}
			application = null;
		}
	}

	internal static void D()
	{
		frmSheetUnhide frmSheetUnhide = new frmSheetUnhide();
		frmSheetUnhide.ShowDialog();
		frmSheetUnhide.Dispose();
		_ = null;
	}
}
