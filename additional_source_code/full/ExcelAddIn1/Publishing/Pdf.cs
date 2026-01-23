using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using ExcelAddIn1.ExcelApp;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.ImportExport;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Publishing;

public sealed class Pdf
{
	public static void ToFolder()
	{
		if (!Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
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
			Microsoft.Office.Interop.Excel.Workbook workbook = null;
			try
			{
				workbook = MH.A.Application.ActiveWorkbook;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (workbook == null)
			{
				return;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				string text;
				try
				{
					text = Path.GetFileNameWithoutExtension(workbook.Name);
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					text = workbook.Name;
					ProjectData.ClearProjectError();
				}
				text += VH.A(97198);
				string text2 = Path.Combine(workbook.Path, text);
				if (!clsPublish.CancelOverwrite(text2))
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
					bool c;
					bool d;
					try
					{
						XmlElement documentElement = KH.A.SettingsXml.DocumentElement;
						c = Conversions.ToBoolean(documentElement.SelectSingleNode(VH.A(97207)).InnerText);
						d = Conversions.ToBoolean(documentElement.SelectSingleNode(VH.A(97244)).InnerText);
						_ = null;
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						Forms.ErrorMessage(VH.A(97285));
						workbook = null;
						ProjectData.ClearProjectError();
						return;
					}
					if (A(workbook, text2, c, d))
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
						clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)4, VH.A(97396));
					}
				}
				workbook = null;
				return;
			}
		}
	}

	private static bool A(Microsoft.Office.Interop.Excel.Workbook A, string B, bool C, bool D)
	{
		bool flag = true;
		List<int> list = new List<int>();
		string text = string.Empty;
		bool flag2 = false;
		int num = 0;
		bool flag3 = false;
		if (EditMode.IsEditMode(A.Application))
		{
			return false;
		}
		Application application = A.Application;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = application.ActiveWindow.SelectedSheets.GetEnumerator();
			while (enumerator.MoveNext())
			{
				object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
				if (objectValue is Worksheet)
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
					Worksheet worksheet = (Worksheet)objectValue;
					if (worksheet.Visible == XlSheetVisibility.xlSheetVisible && Operators.CompareString(worksheet.PageSetup.PrintArea, "", TextCompare: false) != 0)
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
						list.Add(worksheet.Index);
					}
					if (worksheet.ListObjects.Count > 0)
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
						flag3 = true;
					}
					worksheet = null;
				}
				num = checked(num + 1);
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
		if (flag3)
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
			if (num > 1)
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
				Forms.WarningMessage(VH.A(97423));
				goto IL_0577;
			}
		}
		if (list.Any())
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
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			bool copyObjectsWithCells = application.CopyObjectsWithCells;
			application.CopyObjectsWithCells = true;
			Microsoft.Office.Interop.Excel.Workbook workbook;
			try
			{
				workbook = Helpers.A(A, list);
				string listSeparator = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
				IEnumerator enumerator2 = default(IEnumerator);
				Range range;
				try
				{
					enumerator2 = workbook.Worksheets.GetEnumerator();
					bool flag4 = default(bool);
					while (enumerator2.MoveNext())
					{
						Worksheet worksheet2 = (Worksheet)enumerator2.Current;
						if (Operators.CompareString(worksheet2.PageSetup.PrintArea, "", TextCompare: false) == 0)
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
						if (worksheet2.ProtectContents)
						{
							try
							{
								worksheet2.Unprotect(RuntimeHelpers.GetObjectValue(Missing.Value));
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								Forms.ErrorMessage(ex2.Message);
								ProjectData.ClearProjectError();
							}
							flag2 = worksheet2.ProtectContents;
						}
						string[] array = Strings.Split(worksheet2.PageSetup.PrintArea, listSeparator, -1, CompareMethod.Text);
						foreach (string cell in array)
						{
							range = ((_Worksheet)worksheet2).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value));
							if (D)
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
								if (!flag4)
								{
									flag4 = clsImportExport.CheckForErrors(range);
									if (flag4)
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
										Forms.WarningMessage(VH.A(97633));
									}
								}
							}
							if (C && !flag2)
							{
								clsImportExport.RecolorFonts(range, true);
							}
							Range range2 = null;
							if (Operators.ConditionalCompareObjectGreater(range.Cells.CountLarge, 1, TextCompare: false))
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
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									break;
								}
								range2 = range;
							}
							if (range2 == null)
							{
								continue;
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
							Helpers.A(range2, worksheet2, A.Application, A);
						}
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_037c;
						}
						continue;
						end_IL_037c:
						break;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
				range = null;
				try
				{
					application.DisplayAlerts = false;
					((Worksheet)application.ActiveSheet).ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, B, XlFixedFormatQuality.xlQualityStandard, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), true, RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					if (!clsFile.IsPathUrl(B) && !File.Exists(B))
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
						text = ex6.Message;
						clsReporting.LogException(ex6);
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
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					break;
				}
				if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
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
					switch (2)
					{
					case 0:
						continue;
					}
					break;
				}
				if (C)
				{
					Forms.WarningMessage(VH.A(98119));
				}
				else
				{
					Forms.WarningMessage(VH.A(98427));
				}
			}
			workbook = null;
		}
		else
		{
			Forms.WarningMessage(VH.A(98653));
			flag = false;
		}
		goto IL_0577;
		IL_0577:
		application = null;
		return flag;
	}
}
