using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Publishing.Share;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.ImportExport;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Publishing;

public sealed class Send
{
	public static void ShowDialog()
	{
		if (!Access.AllowExcelOperation((PlanType)4, (Restriction)2, false))
		{
			return;
		}
		Application application = MH.A.Application;
		Microsoft.Office.Interop.Excel.Workbook workbook = null;
		string text = string.Empty;
		string text2 = string.Empty;
		try
		{
			workbook = application.ActiveWorkbook;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		wpfSend wpfSend2;
		if (workbook != null)
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
			wpfSend2 = new wpfSend(workbook);
			wpfSend2.ShowDialog();
			if (wpfSend2.DialogResult.HasValue)
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
				if (wpfSend2.DialogResult.Value)
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
					if (wpfSend2.chkSendFile.IsChecked == true)
					{
						string text3 = wpfSend2.txtName.Text;
						if (workbook.Path.Length > 0)
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
							text3 += Path.GetExtension(workbook.Name);
						}
						else
						{
							text3 += VH.A(98730);
						}
						if (wpfSend2.radScopeSelected.IsChecked == true)
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
							text2 = A(workbook, text3);
						}
						else
						{
							text2 = B(workbook, text3);
						}
					}
					if (wpfSend2.chkSendPdf.IsChecked == true)
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
						text = clsPublish.PdfFullName(wpfSend2.txtName.Text + VH.A(97198), workbook.Path, wpfSend2.chkSaveCopy.IsChecked.Value);
						bool? isChecked = wpfSend2.chkSaveCopy.IsChecked;
						if (isChecked.HasValue)
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
							if (isChecked != true)
							{
								goto IL_022b;
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
						}
						if (clsPublish.CancelOverwrite(text))
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
							if (isChecked.HasValue)
							{
								goto IL_031d;
							}
						}
						goto IL_022b;
					}
					goto IL_02b6;
				}
			}
			goto IL_031d;
		}
		goto IL_0322;
		IL_02b6:
		if (wpfSend2.chkSendLink.IsChecked == true)
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
			clsPublish.AttachToEmail(workbook.FullName, true);
		}
		else
		{
			clsPublish.SendAttachment(text, text2, wpfSend2.chkOpen, wpfSend2.chkSaveCopy, wpfSend2.chkCompress);
		}
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)4, VH.A(98741));
		goto IL_031d;
		IL_022b:
		A(workbook, text, wpfSend2.SelectedSheets, wpfSend2.chkCompress.IsChecked.Value, wpfSend2.chkOpen.IsChecked.Value, wpfSend2.chkSaveCopy.IsChecked.Value, wpfSend2.chkRecolor.IsChecked.Value, wpfSend2.chkErrors.IsChecked.Value);
		goto IL_02b6;
		IL_0322:
		application = null;
		return;
		IL_031d:
		wpfSend2 = null;
		workbook = null;
		goto IL_0322;
	}

	private static void A(Microsoft.Office.Interop.Excel.Workbook A, string B, List<int> C, bool D, bool E, bool F, bool G, bool H)
	{
		bool flag = true;
		string text = string.Empty;
		bool flag2 = false;
		Application application = A.Application;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		bool copyObjectsWithCells = application.CopyObjectsWithCells;
		application.CopyObjectsWithCells = true;
		Microsoft.Office.Interop.Excel.Workbook workbook;
		try
		{
			workbook = Helpers.A(A, C);
			string listSeparator = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
			IEnumerator enumerator = default(IEnumerator);
			Range range;
			try
			{
				enumerator = workbook.Worksheets.GetEnumerator();
				bool flag3 = default(bool);
				while (enumerator.MoveNext())
				{
					Worksheet worksheet = (Worksheet)enumerator.Current;
					if (Operators.CompareString(worksheet.PageSetup.PrintArea, "", TextCompare: false) == 0)
					{
						continue;
					}
					if (worksheet.ProtectContents)
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
						try
						{
							worksheet.Unprotect(RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							Forms.ErrorMessage(ex2.Message);
							ProjectData.ClearProjectError();
						}
						flag2 = worksheet.ProtectContents;
					}
					string[] array = Strings.Split(worksheet.PageSetup.PrintArea, listSeparator, -1, CompareMethod.Text);
					foreach (string cell in array)
					{
						range = ((_Worksheet)worksheet).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value));
						if (H)
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
							if (!flag3)
							{
								flag3 = clsImportExport.CheckForErrors(range);
								if (flag3)
								{
									Forms.WarningMessage(VH.A(97633));
								}
							}
						}
						if (G && !flag2)
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
							clsImportExport.RecolorFonts(range, true);
						}
						Range range2 = null;
						if (Operators.ConditionalCompareObjectGreater(range.Cells.CountLarge, 1, TextCompare: false))
						{
							try
							{
								range2 = range.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value));
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
						}
						else if (Conversions.ToBoolean(range.HasFormula))
						{
							range2 = range;
						}
						if (range2 != null)
						{
							Helpers.A(range2, worksheet, A.Application, A);
						}
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_0211;
					}
					continue;
					end_IL_0211:
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			range = null;
			try
			{
				application.DisplayAlerts = false;
				((Worksheet)application.ActiveSheet).ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, B, XlFixedFormatQuality.xlQualityStandard, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), E, RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				if (!clsFile.IsPathUrl(B))
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
					if (!File.Exists(B))
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
						text = ex6.Message;
						clsReporting.LogException(ex6);
					}
				}
				flag = false;
				ProjectData.ClearProjectError();
			}
			finally
			{
				application.DisplayAlerts = true;
			}
			workbook.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			text = ex8.Message;
			flag = false;
			clsReporting.LogException(ex8);
			ProjectData.ClearProjectError();
		}
		application.CopyObjectsWithCells = copyObjectsWithCells;
		application.ScreenUpdating = true;
		application.EnableEvents = true;
		if (!flag)
		{
			if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
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
				Forms.ErrorMessage(VH.A(97793) + Path.GetFileName(B) + VH.A(97830));
			}
			else
			{
				Forms.ErrorMessage(VH.A(97998) + text);
			}
		}
		else if (flag2)
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
			if (G)
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
				Forms.WarningMessage(VH.A(98119));
			}
			else
			{
				Forms.WarningMessage(VH.A(98427));
			}
		}
		workbook = null;
		application = null;
	}

	private static string A(Microsoft.Office.Interop.Excel.Workbook A, string B)
	{
		Application application = A.Application;
		Window activeWindow = application.ActiveWindow;
		string result = string.Empty;
		Application application2 = application;
		application2.DisplayAlerts = false;
		application2.EnableEvents = false;
		application2.ScreenUpdating = false;
		bool copyObjectsWithCells = application2.CopyObjectsWithCells;
		application2.CopyObjectsWithCells = true;
		try
		{
			Microsoft.Office.Interop.Excel.Workbook workbook = Helpers.A(A, application);
			Microsoft.Office.Interop.Excel.Workbook workbook2 = workbook;
			for (int i = workbook2.Worksheets.Count; i >= 2; i = checked(i + -1))
			{
				NewLateBinding.LateCall(workbook2.Worksheets[i], null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
			}
			activeWindow.SelectedSheets.Copy(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(workbook2.Worksheets[1]));
			((Worksheet)workbook2.Worksheets[1]).Delete();
			workbook2 = null;
			bool num = Helpers.A(workbook);
			if (!num)
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
				workbook.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			if (num)
			{
				string text = Base.SaveTempWorkbook(A, workbook);
				I.A.FileSystem.RenameFile(text, B);
				try
				{
					result = Path.Combine(Path.GetDirectoryName(text), B);
				}
				catch (ArgumentException ex)
				{
					ProjectData.SetProjectError(ex);
					ArgumentException ex2 = ex;
					Forms.ErrorMessage(ex2.Message);
					ProjectData.ClearProjectError();
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.ErrorMessage(ex4.Message);
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
		finally
		{
			Microsoft.Office.Interop.Excel.Workbook workbook = null;
		}
		application2.CopyObjectsWithCells = copyObjectsWithCells;
		application2.DisplayAlerts = true;
		application2.EnableEvents = true;
		application2.ScreenUpdating = true;
		application2 = null;
		application = null;
		activeWindow = null;
		return result;
	}

	private static string B(Microsoft.Office.Interop.Excel.Workbook A, string B)
	{
		if (Operators.CompareString(A.Name, B, TextCompare: false) == 0)
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
					return A.FullName;
				}
			}
		}
		string text = Path.Combine(I.A.FileSystem.SpecialDirectories.Temp, B);
		A.SaveCopyAs(text);
		return text;
	}
}
