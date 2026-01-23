using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Sheets;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Comments;

public sealed class CleanUp
{
	public static void Delete()
	{
		bool flag = false;
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		if (application.Selection is Range)
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
			Range range = (Range)application.Selection;
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			int num = default(int);
			int num2 = default(int);
			int num3 = default(int);
			Microsoft.Office.Interop.Excel.Sheets selectedSheets;
			object objectValue;
			try
			{
				selectedSheets = application.ActiveWindow.SelectedSheets;
				if (selectedSheets.Count > 1)
				{
					if (MessageBox.Show(VH.A(141780), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
					{
						IEnumerator enumerator = default(IEnumerator);
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							num = 0;
							num2 = 0;
							num3 = 0;
							objectValue = RuntimeHelpers.GetObjectValue(application.ActiveSheet);
							Worksheet worksheet;
							try
							{
								enumerator = selectedSheets.GetEnumerator();
								while (enumerator.MoveNext())
								{
									object objectValue2 = RuntimeHelpers.GetObjectValue(enumerator.Current);
									if (!(objectValue2 is Worksheet))
									{
										continue;
									}
									while (true)
									{
										switch (3)
										{
										case 0:
											continue;
										}
										break;
									}
									worksheet = (Worksheet)objectValue2;
									if (!worksheet.ProtectContents)
									{
										try
										{
											worksheet.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
											num2 = Conversions.ToInteger(Operators.AddObject(num2, worksheet.Cells.SpecialCells(XlCellType.xlCellTypeComments, RuntimeHelpers.GetObjectValue(Missing.Value)).Cells.CountLarge));
										}
										catch (Exception ex)
										{
											ProjectData.SetProjectError(ex);
											Exception ex2 = ex;
											ProjectData.ClearProjectError();
										}
										try
										{
											num = Conversions.ToInteger(Operators.AddObject(num, NewLateBinding.LateGet(NewLateBinding.LateGet(worksheet, null, VH.A(141942), new object[0], null, null, null), null, VH.A(52690), new object[0], null, null, null)));
										}
										catch (Exception ex3)
										{
											ProjectData.SetProjectError(ex3);
											Exception ex4 = ex3;
											ProjectData.ClearProjectError();
										}
										worksheet.Cells.ClearComments();
									}
									num3 = checked(num3 + 1);
								}
							}
							finally
							{
								if (enumerator is IDisposable)
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											continue;
										}
										(enumerator as IDisposable).Dispose();
										break;
									}
								}
							}
							worksheet = null;
							selectedSheets.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
							NewLateBinding.LateCall(objectValue, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
							range.Select();
							flag = true;
							break;
						}
					}
				}
				else
				{
					ExcelAddIn1.Sheets.Protection.Unprotect(range.Worksheet);
					if (!range.Worksheet.ProtectContents && MessageBox.Show(VH.A(141975), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							range.ClearComments();
							break;
						}
					}
				}
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			application.EnableEvents = true;
			application.ScreenUpdating = true;
			if (flag)
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
				Forms.InfoMessage(VH.A(52374) + num + VH.A(142125) + num2 + VH.A(142154) + num3 + VH.A(142175));
			}
			range = null;
			selectedSheets = null;
			objectValue = null;
		}
		application = null;
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(142200));
	}

	public static void DeleteEmptyNotes()
	{
		Regex regex = new Regex(VH.A(142247));
		if (MessageBox.Show(VH.A(142288), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			application.ScreenUpdating = false;
			int num = default(int);
			try
			{
				int count = application.ActiveWindow.SelectedSheets.Count;
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = application.ActiveWorkbook.Worksheets.GetEnumerator();
					IEnumerator enumerator2 = default(IEnumerator);
					while (enumerator.MoveNext())
					{
						Worksheet worksheet = (Worksheet)enumerator.Current;
						if (count == 1)
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
							ExcelAddIn1.Sheets.Protection.Unprotect(worksheet);
						}
						if (worksheet.ProtectContents)
						{
							continue;
						}
						try
						{
							enumerator2 = worksheet.Comments.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Comment comment = (Comment)enumerator2.Current;
								if (!regex.IsMatch(comment.Text(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value))))
								{
									continue;
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									break;
								}
								comment.Delete();
								num = checked(num + 1);
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_013e;
								}
								continue;
								end_IL_013e:
								break;
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									(enumerator2 as IDisposable).Dispose();
									break;
								}
							}
						}
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_0178;
						}
						continue;
						end_IL_0178:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (1)
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
				ProjectData.ClearProjectError();
			}
			application.ScreenUpdating = true;
			application = null;
			string text;
			if (num == 1)
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
				text = VH.A(52374) + Conversions.ToString(num) + VH.A(142472);
			}
			else
			{
				text = VH.A(52374) + Conversions.ToString(num) + VH.A(142497);
			}
			Forms.InfoMessage(text);
		}
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(142524));
	}
}
