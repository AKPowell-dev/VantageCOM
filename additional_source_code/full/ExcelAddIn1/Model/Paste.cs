using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Audit.Visualizations;
using ExcelAddIn1.Format;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Model;

public sealed class Paste
{
	[CompilerGenerated]
	private static Range m_A;

	[CompilerGenerated]
	private static Range m_B;

	public static Range CopiedRange
	{
		[CompilerGenerated]
		get
		{
			return Paste.m_A;
		}
		[CompilerGenerated]
		set
		{
			Paste.m_A = value;
		}
	} = null;

	public static Range CutRange
	{
		[CompilerGenerated]
		get
		{
			return Paste.m_B;
		}
		[CompilerGenerated]
		set
		{
			Paste.m_B = value;
		}
	} = null;

	public static void Copy()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
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
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				CopiedRange = (Range)application.Selection;
			}
			CutRange = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application = null;
	}

	public static void Cut()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		try
		{
			if (application.Selection is Range)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				CutRange = (Range)application.Selection;
			}
			CopiedRange = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application = null;
	}

	public static void NumberFormats(bool blnTranspose = false)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		Range range2 = default(Range);
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
			if (CopiedRange == null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						B();
						return;
					}
				}
			}
			Microsoft.Office.Interop.Excel.Application application = CopiedRange.Application;
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			Range range;
			Range range3;
			try
			{
				range = (Range)((Range)application.Selection).Cells[1, 1];
				Range copiedRange = CopiedRange;
				int num = Conversions.ToInteger(copiedRange.Rows.CountLarge);
				int num2 = Conversions.ToInteger(copiedRange.Columns.CountLarge);
				if (blnTranspose)
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
					range2 = range.get_Resize((object)num2, (object)num);
				}
				else
				{
					range2 = (Range)application.Selection;
					int num3 = Conversions.ToInteger(range2.Rows.CountLarge);
					int num4 = Conversions.ToInteger(range2.Columns.CountLarge);
					if (num3 % num == 0)
					{
						if (num4 % num2 == 0)
						{
							goto IL_013d;
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
					range2 = range.get_Resize((object)num, (object)num2);
				}
				goto IL_013d;
				IL_013d:
				bool flag = JH.A(range2);
				if (blnTranspose)
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
					try
					{
						enumerator = copiedRange.Cells.GetEnumerator();
						while (enumerator.MoveNext())
						{
							range3 = (Range)enumerator.Current;
							int num5 = Conversions.ToInteger(Operators.SubtractObject(range3.Row, NewLateBinding.LateGet(copiedRange.Cells[1, 1], null, VH.A(91823), new object[0], null, null, null)));
							int num6 = Conversions.ToInteger(Operators.SubtractObject(range3.Column, NewLateBinding.LateGet(copiedRange.Cells[1, 1], null, VH.A(41354), new object[0], null, null, null)));
							range.get_Offset((object)num6, (object)num5).NumberFormat = RuntimeHelpers.GetObjectValue(range3.NumberFormat);
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
				else
				{
					object objectValue;
					try
					{
						objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range2, null, VH.A(1998), new object[0], null, null, null));
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						objectValue = RuntimeHelpers.GetObjectValue(range2.Formula);
						ProjectData.ClearProjectError();
					}
					range2.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					try
					{
						NewLateBinding.LateSet(range2, null, VH.A(1998), new object[1] { objectValue }, null, null);
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						range2.Formula = RuntimeHelpers.GetObjectValue(objectValue);
						ProjectData.ClearProjectError();
					}
					objectValue = null;
				}
				copiedRange = null;
				range2.Select();
				if (flag)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						JH.A(range2, VH.A(91830));
						break;
					}
				}
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			try
			{
				C(range2);
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
			application.ScreenUpdating = true;
			application.EnableEvents = true;
			application = null;
			range2 = null;
			range = null;
			range3 = null;
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(91869));
			return;
		}
	}

	public static void Duplicate()
	{
		if (!Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		bool flag = default(bool);
		Range range = default(Range);
		Range range2 = default(Range);
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
			Microsoft.Office.Interop.Excel.Window window = null;
			Microsoft.Office.Interop.Excel.Workbook workbook = null;
			if (CopiedRange == null)
			{
				B();
				return;
			}
			try
			{
				if (CopiedRange.Areas.Count > 1)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							C();
							return;
						}
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				CopiedRange = null;
				ProjectData.ClearProjectError();
				return;
			}
			try
			{
				workbook = MH.A.Application.ActiveWorkbook;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			if (workbook == null)
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
				if (Workbooks.IsShared(workbook, true, (System.Windows.Window)null))
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							workbook = null;
							return;
						}
					}
				}
				workbook = null;
				Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
				try
				{
					workbook = application.ActiveWorkbook;
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
				if (workbook != null)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
						{
							if (Workbooks.IsShared(workbook, true, (System.Windows.Window)null))
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										break;
									default:
										workbook = null;
										application = null;
										return;
									}
								}
							}
							if (workbook.Windows.Count > 1)
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
								window = application.ActiveWindow;
							}
							workbook = null;
							Microsoft.Office.Interop.Excel.Application application2 = application;
							application2.ScreenUpdating = false;
							application2.EnableEvents = false;
							application2.DisplayAlerts = false;
							bool displayStatusBar = application2.DisplayStatusBar;
							XlCalculation calculation = application2.Calculation;
							application2.Calculation = XlCalculation.xlCalculationManual;
							_ = null;
							string cell = CopiedRange.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							try
							{
								range = application.ActiveCell;
								Range copiedRange = CopiedRange;
								copiedRange.Worksheet.Copy(RuntimeHelpers.GetObjectValue(application.ActiveSheet), RuntimeHelpers.GetObjectValue(Missing.Value));
								Worksheet obj = (Worksheet)application.ActiveSheet;
								range2 = ((_Worksheet)obj).get_Range((object)copiedRange.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value));
								int num = Conversions.ToInteger(copiedRange.Rows.CountLarge);
								int num2 = Conversions.ToInteger(copiedRange.Columns.CountLarge);
								copiedRange = null;
								string name = range.Worksheet.Name;
								string name2 = CopiedRange.Worksheet.Name;
								string name3 = obj.Name;
								Range obj2 = (Range)application.Selection;
								int num3 = Conversions.ToInteger(obj2.Rows.CountLarge);
								int num4 = Conversions.ToInteger(obj2.Columns.CountLarge);
								if (num3 % num == 0)
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
									if (num4 % num2 == 0)
									{
										goto IL_0311;
									}
								}
								range.get_Resize((object)num, (object)num2);
								goto IL_0311;
								IL_0311:
								range = range.get_Resize((object)num, (object)num2);
								try
								{
									num3 = checked((int)Math.Round(application.WorksheetFunction.CountA(range, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value))));
								}
								catch (Exception ex7)
								{
									ProjectData.SetProjectError(ex7);
									Exception ex8 = ex7;
									ProjectData.ClearProjectError();
								}
								if (num3 > 0 && System.Windows.Forms.MessageBox.Show(VH.A(91910), VH.A(40448), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											throw new Exception();
										}
									}
								}
								flag = JH.A(range);
								range2.Cut(RuntimeHelpers.GetObjectValue(Missing.Value));
								if (window != null)
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
									window.Activate();
									window = null;
								}
								range.Worksheet.Activate();
								if (!(application.Selection is Range))
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
									range.Select();
								}
								try
								{
									range.Worksheet.Paste(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								}
								catch (Exception ex9)
								{
									ProjectData.SetProjectError(ex9);
									Exception ex10 = ex9;
									if (ex10.Message.Contains(VH.A(92192)))
									{
										Forms.ErrorMessage(VH.A(92229));
										clsReporting.LogException(ex10);
									}
									throw;
								}
								CopiedRange = ((_Worksheet)CopiedRange.Worksheet).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value));
								try
								{
									if (Operators.CompareString(name, name2, TextCompare: false) == 0)
									{
										while (true)
										{
											switch (5)
											{
											case 0:
												break;
											default:
												A(range, name3, "");
												goto end_IL_05dc;
											}
											continue;
											end_IL_05dc:
											break;
										}
									}
									else
									{
										A(range, name3, name2);
									}
								}
								catch (Exception ex11)
								{
									ProjectData.SetProjectError(ex11);
									Exception ex12 = ex11;
									ProjectData.ClearProjectError();
								}
							}
							catch (Exception ex13)
							{
								ProjectData.SetProjectError(ex13);
								Exception ex14 = ex13;
								ProjectData.ClearProjectError();
							}
							finally
							{
								try
								{
									range2.Worksheet.Delete();
								}
								catch (Exception ex15)
								{
									ProjectData.SetProjectError(ex15);
									Exception ex16 = ex15;
									ProjectData.ClearProjectError();
								}
							}
							try
							{
								if (flag)
								{
									JH.A(range, VH.A(92805));
								}
								C(range);
							}
							catch (Exception ex17)
							{
								ProjectData.SetProjectError(ex17);
								Exception ex18 = ex17;
								ProjectData.ClearProjectError();
							}
							Microsoft.Office.Interop.Excel.Application application3 = application;
							application3.DisplayStatusBar = displayStatusBar;
							application3.Calculation = calculation;
							CopiedRange.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
							Common.RefreshLiveVisualizations(range);
							application3.DisplayAlerts = true;
							application3.ScreenUpdating = true;
							application3.EnableEvents = true;
							_ = null;
							application = null;
							range = null;
							range2 = null;
							clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(92805));
							return;
						}
						}
					}
				}
				application = null;
				return;
			}
		}
	}

	private static void A(Range A, string B, string C)
	{
		B = B.Replace(VH.A(39851), VH.A(39854));
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range range = (Range)enumerator.Current;
				if (Paste.A(range))
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				string text;
				try
				{
					text = Conversions.ToString(NewLateBinding.LateGet(range, null, VH.A(1998), new object[0], null, null, null));
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					text = Conversions.ToString(range.Formula);
					ProjectData.ClearProjectError();
				}
				if (text.Contains(B + VH.A(7827)) | text.Contains(B + VH.A(43343)))
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
					if (C.Length == 0)
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
						text = text.Replace(VH.A(39851) + B + VH.A(43343), C);
					}
					else
					{
						text = text.Replace(B, C.Replace(VH.A(39851), VH.A(39854)));
					}
				}
				try
				{
					NewLateBinding.LateSet(range, null, VH.A(1998), new object[1] { text }, null, null);
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					range.Formula = text;
					ProjectData.ClearProjectError();
				}
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (7)
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

	public static void Exact(bool trans)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		if (CopiedRange == null)
		{
			B();
			return;
		}
		try
		{
			if (CopiedRange.Areas.Count > 1)
			{
				C();
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			CopiedRange = null;
			ProjectData.ClearProjectError();
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = CopiedRange.Application;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		Range range2 = default(Range);
		Range range = default(Range);
		Range range3;
		try
		{
			range = (Range)((Range)application.Selection).Cells[1, 1];
			Range copiedRange = CopiedRange;
			copiedRange.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
			int num = Conversions.ToInteger(copiedRange.Rows.CountLarge);
			int num2 = Conversions.ToInteger(copiedRange.Columns.CountLarge);
			bool flag;
			if (trans)
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
				range2 = range.get_Resize((object)num2, (object)num);
				flag = JH.A(range2);
				range2.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, true);
				goto IL_0238;
			}
			range2 = (Range)application.Selection;
			int num3 = Conversions.ToInteger(range2.Rows.CountLarge);
			int num4 = Conversions.ToInteger(range2.Columns.CountLarge);
			if (num3 % num == 0)
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
				if (num4 % num2 == 0)
				{
					goto IL_01b2;
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
			range2 = range.get_Resize((object)num, (object)num2);
			goto IL_01b2;
			IL_01b2:
			flag = JH.A(range2);
			try
			{
				range.Worksheet.Paste(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				if (ex4.Message.Contains(VH.A(92192)))
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
					Forms.ErrorMessage(VH.A(92229));
					clsReporting.LogException(ex4);
				}
				throw;
			}
			range2 = (Range)application.Selection;
			goto IL_0238;
			IL_0238:
			if (trans)
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
				{
					IEnumerator enumerator = copiedRange.Cells.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							range3 = (Range)enumerator.Current;
							int num5 = Conversions.ToInteger(Operators.SubtractObject(range3.Row, NewLateBinding.LateGet(copiedRange.Cells[1, 1], null, VH.A(41347), new object[0], null, null, null)));
							int num6 = Conversions.ToInteger(Operators.SubtractObject(range3.Column, NewLateBinding.LateGet(copiedRange.Cells[1, 1], null, VH.A(41354), new object[0], null, null, null)));
							try
							{
								NewLateBinding.LateSetComplex(range.get_Offset((object)num6, (object)num5), null, VH.A(1998), new object[1] { NewLateBinding.LateGet(range3, null, VH.A(1998), new object[0], null, null, null) }, null, null, OptimisticSet: false, RValueBase: true);
							}
							catch (Exception ex5)
							{
								ProjectData.SetProjectError(ex5);
								Exception ex6 = ex5;
								range.get_Offset((object)num6, (object)num5).Formula = RuntimeHelpers.GetObjectValue(range3.Formula);
								ProjectData.ClearProjectError();
							}
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_03a3;
							}
							continue;
							end_IL_03a3:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
			}
			else
			{
				num3 = Conversions.ToInteger(range2.Rows.CountLarge);
				num4 = Conversions.ToInteger(range2.Columns.CountLarge);
				if (num3 == num)
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
					if (num4 == num2)
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
						try
						{
							NewLateBinding.LateSet(range2, null, VH.A(1998), new object[1] { NewLateBinding.LateGet(copiedRange, null, VH.A(1998), new object[0], null, null, null) }, null, null);
						}
						catch (Exception ex7)
						{
							ProjectData.SetProjectError(ex7);
							Exception ex8 = ex7;
							range2.Formula = RuntimeHelpers.GetObjectValue(copiedRange.Formula);
							ProjectData.ClearProjectError();
						}
						try
						{
							if (Operators.ConditionalCompareObjectEqual(copiedRange.Cells.CountLarge, 1, TextCompare: false))
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									if (A(CopiedRange))
									{
										range2.Value2 = Operators.ConcatenateObject(VH.A(39851), copiedRange.Value2);
									}
									break;
								}
							}
							else
							{
								Range range4 = null;
								Range range5 = null;
								try
								{
									range4 = copiedRange.SpecialCells(XlCellType.xlCellTypeConstants, RuntimeHelpers.GetObjectValue(Missing.Value));
								}
								catch (Exception ex9)
								{
									ProjectData.SetProjectError(ex9);
									Exception ex10 = ex9;
									ProjectData.ClearProjectError();
								}
								if (range4 != null)
								{
									while (true)
									{
										switch (5)
										{
										case 0:
											continue;
										}
										try
										{
											range5 = range2.SpecialCells(XlCellType.xlCellTypeConstants, RuntimeHelpers.GetObjectValue(Missing.Value));
										}
										catch (Exception ex11)
										{
											ProjectData.SetProjectError(ex11);
											Exception ex12 = ex11;
											ProjectData.ClearProjectError();
										}
										int num7 = Conversions.ToInteger(range4.Cells.CountLarge);
										if (range5 != null)
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
											if (!Operators.ConditionalCompareObjectLess(range5.Cells.CountLarge, num7, TextCompare: false))
											{
												if (num7 <= 25)
												{
													B(CopiedRange, range4, range2);
												}
												goto IL_05ce;
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
										}
										if (num7 <= 25)
										{
											B(CopiedRange, range4, range2);
										}
										else
										{
											Forms.WarningMessage(VH.A(92836));
										}
										goto IL_05ce;
										IL_05ce:
										range4 = null;
										range5 = null;
										break;
									}
								}
							}
						}
						catch (Exception ex13)
						{
							ProjectData.SetProjectError(ex13);
							Exception ex14 = ex13;
							clsReporting.LogException(ex14);
							ProjectData.ClearProjectError();
						}
						goto IL_0934;
					}
				}
				if (num3 % num == 0)
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
					if (num4 % num2 == 0)
					{
						Range range6 = null;
						try
						{
							range6 = range2.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						catch (Exception ex15)
						{
							ProjectData.SetProjectError(ex15);
							Exception ex16 = ex15;
							ProjectData.ClearProjectError();
						}
						if (range6 != null)
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
							int num8 = Conversions.ToInteger(NewLateBinding.LateGet(range2.Cells[1, 1], null, VH.A(41347), new object[0], null, null, null));
							int num9 = Conversions.ToInteger(NewLateBinding.LateGet(range2.Cells[1, 1], null, VH.A(41354), new object[0], null, null, null));
							{
								IEnumerator enumerator2 = range6.GetEnumerator();
								try
								{
									while (enumerator2.MoveNext())
									{
										Range range7 = (Range)enumerator2.Current;
										int num5 = checked(range7.Row - num8) % num;
										int num6 = checked(range7.Column - num9) % num2;
										try
										{
											string memberName = VH.A(1998);
											object[] array = new object[1];
											object instance = CopiedRange.Cells[1, 1];
											string memberName2 = VH.A(60565);
											object[] obj = new object[2] { num5, num6 };
											object[] array2 = obj;
											bool[] obj2 = new bool[2] { true, true };
											bool[] array3 = obj2;
											object instance2 = NewLateBinding.LateGet(instance, null, memberName2, obj, null, null, obj2);
											if (array3[0])
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
												num5 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
											}
											if (array3[1])
											{
												num6 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[1]), typeof(int));
											}
											array[0] = NewLateBinding.LateGet(instance2, null, VH.A(1998), new object[0], null, null, null);
											NewLateBinding.LateSet(range7, null, memberName, array, null, null);
										}
										catch (Exception ex17)
										{
											ProjectData.SetProjectError(ex17);
											Exception ex18 = ex17;
											object instance3 = CopiedRange.Cells[1, 1];
											string memberName3 = VH.A(60565);
											object[] obj3 = new object[2] { num5, num6 };
											object[] array2 = obj3;
											bool[] obj4 = new bool[2] { true, true };
											bool[] array3 = obj4;
											object instance4 = NewLateBinding.LateGet(instance3, null, memberName3, obj3, null, null, obj4);
											if (array3[0])
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
												num5 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
											}
											if (array3[1])
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
												num6 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[1]), typeof(int));
											}
											range7.Formula = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance4, null, VH.A(68956), new object[0], null, null, null));
											ProjectData.ClearProjectError();
										}
									}
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											goto end_IL_0912;
										}
										continue;
										end_IL_0912:
										break;
									}
								}
								finally
								{
									IDisposable disposable2 = enumerator2 as IDisposable;
									if (disposable2 != null)
									{
										disposable2.Dispose();
									}
								}
							}
						}
					}
				}
			}
			goto IL_0934;
			IL_0934:
			copiedRange = null;
			range2.Select();
			if (flag)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					JH.A(range2, VH.A(93062));
					break;
				}
			}
		}
		catch (Exception ex19)
		{
			ProjectData.SetProjectError(ex19);
			Exception ex20 = ex19;
			ProjectData.ClearProjectError();
		}
		C(range2);
		Common.RefreshLiveVisualizations(range);
		application.ScreenUpdating = true;
		application.EnableEvents = true;
		application = null;
		range2 = null;
		range = null;
		range3 = null;
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(93062));
	}

	private static bool A(Range A)
	{
		return Operators.CompareString(A.PrefixCharacter.ToString(), VH.A(39851), TextCompare: false) == 0;
	}

	private static void B(Range A, Range B, Range C)
	{
		Range obj = (Range)A.Cells[1, 1];
		int row = obj.Row;
		int column = obj.Column;
		_ = null;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = B.Cells.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range range = (Range)enumerator.Current;
				if (!Paste.A(range))
				{
					continue;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Range range2 = range;
				checked(((Range)C.Cells[1, 1]).get_Offset((object)(range2.Row - row), (object)(range2.Column - column))).Value2 = Operators.ConcatenateObject(VH.A(39851), range2.Value2);
				range2 = null;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					return;
				}
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
	}

	public static void Links(bool blnTranspose = false)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		if (CopiedRange == null)
		{
			B();
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = CopiedRange.Application;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		Range range2 = default(Range);
		Range range = default(Range);
		Range range3;
		try
		{
			range = (Range)((Range)application.Selection).Cells[1, 1];
			Range copiedRange = CopiedRange;
			int num = Conversions.ToInteger(copiedRange.Rows.CountLarge);
			int num2 = Conversions.ToInteger(copiedRange.Columns.CountLarge);
			if (blnTranspose)
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
				range2 = range.get_Resize((object)num2, (object)num);
			}
			else
			{
				range2 = (Range)application.Selection;
				int num3 = Conversions.ToInteger(range2.Rows.CountLarge);
				int num4 = Conversions.ToInteger(range2.Columns.CountLarge);
				if (num3 % num == 0)
				{
					if (num4 % num2 == 0)
					{
						goto IL_0133;
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
				}
				range2 = range.get_Resize((object)num, (object)num2);
			}
			goto IL_0133;
			IL_0133:
			bool flag = JH.A(range2);
			if (blnTranspose)
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
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = copiedRange.Cells.GetEnumerator();
					while (enumerator.MoveNext())
					{
						range3 = (Range)enumerator.Current;
						if (!string.IsNullOrEmpty(Conversions.ToString(range3.Text)))
						{
							int num5 = Conversions.ToInteger(Operators.SubtractObject(range3.Row, NewLateBinding.LateGet(copiedRange.Cells[1, 1], null, VH.A(91823), new object[0], null, null, null)));
							int num6 = Conversions.ToInteger(Operators.SubtractObject(range3.Column, NewLateBinding.LateGet(copiedRange.Cells[1, 1], null, VH.A(41354), new object[0], null, null, null)));
							A(range.get_Offset((object)num6, (object)num5), range3);
						}
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_0252;
						}
						continue;
						end_IL_0252:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (6)
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
			else
			{
				try
				{
					try
					{
						range.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						range.Worksheet.Paste(RuntimeHelpers.GetObjectValue(Missing.Value), true);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						if (ex2.Message.Contains(VH.A(92192)))
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
							Forms.ErrorMessage(VH.A(92229));
							clsReporting.LogException(ex2);
						}
						throw;
					}
					IEnumerator enumerator2 = range2.GetEnumerator();
					try
					{
						while (enumerator2.MoveNext())
						{
							Range range4 = (Range)enumerator2.Current;
							object objectValue;
							try
							{
								objectValue = RuntimeHelpers.GetObjectValue(application.Evaluate(RuntimeHelpers.GetObjectValue(range4.Formula)));
								if (objectValue is Range)
								{
									while (true)
									{
										switch (2)
										{
										case 0:
											continue;
										}
										if (!string.IsNullOrEmpty(Conversions.ToString(((Range)objectValue).Text)))
										{
											break;
										}
										while (true)
										{
											switch (4)
											{
											case 0:
												continue;
											}
											range4.Formula = "";
											break;
										}
										break;
									}
								}
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
							objectValue = null;
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_03bc;
							}
							continue;
							end_IL_03bc:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator2 as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					if (CopiedRange.Areas.Count > 1)
					{
						Forms.WarningMessage(VH.A(93085));
					}
					else
					{
						Forms.ErrorMessage(VH.A(93303) + ex6.Message);
						clsReporting.LogException(ex6);
					}
					throw ex6;
				}
			}
			copiedRange = null;
			try
			{
				if (KH.A.AutoColorOnEntry)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						AutoColor.Range(range2);
						break;
					}
				}
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
			range2.Select();
			if (flag)
			{
				JH.A(range2, VH.A(93408));
			}
		}
		catch (Exception ex9)
		{
			ProjectData.SetProjectError(ex9);
			Exception ex10 = ex9;
			ProjectData.ClearProjectError();
		}
		C(range2);
		Common.RefreshLiveVisualizations(range);
		application.ScreenUpdating = true;
		application.EnableEvents = true;
		application = null;
		range2 = null;
		range = null;
		range3 = null;
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(93408));
	}

	private static void A(Range A, Range B)
	{
		Range range = A;
		if (range.Worksheet == B.Worksheet)
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
			range.Formula = VH.A(48936) + B.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		else
		{
			range.Formula = VH.A(48936) + B.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		range.NumberFormat = RuntimeHelpers.GetObjectValue(B.NumberFormat);
		range = null;
	}

	private static void C(Range A)
	{
		try
		{
			if (A.Application.Intersect(CopiedRange, A, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == null)
			{
				CopiedRange.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			else
			{
				CopiedRange = null;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void Insert()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		checked
		{
			XlCalculation calculation = default(XlCalculation);
			bool flag2 = default(bool);
			bool flag3 = default(bool);
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				bool flag;
				Range range;
				if (CopiedRange != null)
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
					range = CopiedRange;
					flag = false;
				}
				else
				{
					if (CutRange == null)
					{
						B();
						return;
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
					range = CutRange;
					flag = true;
				}
				Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
				Common.ClearVisualizations(application);
				int num;
				int num2;
				int column;
				int row;
				Range activeCell;
				Range range2;
				try
				{
					Microsoft.Office.Interop.Excel.Application application2 = application;
					application2.ScreenUpdating = false;
					application2.EnableEvents = false;
					calculation = application2.Calculation;
					application2.Calculation = XlCalculation.xlCalculationManual;
					num = Conversions.ToInteger(range.Rows.CountLarge);
					num2 = Conversions.ToInteger(range.Columns.CountLarge);
					if (Operators.ConditionalCompareObjectEqual(num2, application2.Columns.CountLarge, TextCompare: false))
					{
						flag2 = true;
					}
					if (Operators.ConditionalCompareObjectEqual(num, application2.Rows.CountLarge, TextCompare: false))
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
						flag3 = true;
					}
					range2 = (Range)application2.Selection;
					activeCell = application2.ActiveCell;
					application2 = null;
					column = activeCell.Column;
					row = activeCell.Row;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					Forms.ErrorMessage(VH.A(93431) + ex2.Message);
					ProjectData.ClearProjectError();
					goto IL_058b;
				}
				if (flag)
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
					range.Cut(RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				else
				{
					range.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				Range range4;
				try
				{
					Range range3 = activeCell;
					if (flag2)
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
						range3.get_Resize((object)num, RuntimeHelpers.GetObjectValue(Missing.Value)).EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
						range4 = (Range)range3.Worksheet.Cells[range3.get_Offset((object)(-num), (object)0).Row, 1];
						goto IL_042d;
					}
					if (flag3)
					{
						range3.get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)num2).EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight, RuntimeHelpers.GetObjectValue(Missing.Value));
						range4 = (Range)range3.Worksheet.Cells[1, range3.get_Offset((object)0, (object)(-num2)).Column];
						goto IL_042d;
					}
					wpfPasteInsert wpfPasteInsert2 = new wpfPasteInsert();
					if (wpfPasteInsert2.ShowDialog() == true)
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
						bool pasteRows = wpfPasteInsert2.PasteRows;
						wpfPasteInsert2 = null;
						if (pasteRows)
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
							range2.get_Resize((object)num, (object)num2).Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
							range4 = (Range)range3.Worksheet.Cells[range3.get_Offset((object)(-num), (object)0).Row, range3.Column];
						}
						else
						{
							range2.get_Resize((object)num, (object)num2).Insert(XlInsertShiftDirection.xlShiftToRight, RuntimeHelpers.GetObjectValue(Missing.Value));
							range4 = (Range)range3.Worksheet.Cells[range3.Row, range3.get_Offset((object)0, (object)(-num2)).Column];
						}
						goto IL_042d;
					}
					goto end_IL_01b5;
					IL_042d:
					range3 = null;
					try
					{
						Range range5 = range4;
						if (flag2)
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
							range5.EntireRow.get_Resize((object)num, RuntimeHelpers.GetObjectValue(Missing.Value)).Select();
							NewLateBinding.LateCall(range5.Columns[column, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
						}
						else if (flag3)
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
							range5.EntireColumn.get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)num2).Select();
							NewLateBinding.LateCall(range5.Rows[row, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
						}
						else
						{
							range5.get_Resize((object)num, (object)num2).Select();
						}
						range5 = null;
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
					end_IL_01b5:;
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					Forms.ErrorMessage(VH.A(93484) + ex6.Message);
					ProjectData.ClearProjectError();
				}
				goto IL_058b;
				IL_058b:
				Microsoft.Office.Interop.Excel.Application application3 = application;
				application3.Calculation = calculation;
				if (flag)
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
					application3.CutCopyMode = (XlCutCopyMode)0;
					CutRange = null;
				}
				else
				{
					CopiedRange.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				application3.ScreenUpdating = true;
				application3.EnableEvents = true;
				application3 = null;
				application = null;
				range4 = null;
				activeCell = null;
				range2 = null;
				range = null;
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(93537));
				return;
			}
		}
	}

	public static void Transpose()
	{
		if (CopiedRange != null)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					wpfPasteTrans wpfPasteTrans2 = new wpfPasteTrans();
					if (wpfPasteTrans2.ShowDialog() == true)
					{
						try
						{
							if (wpfPasteTrans2.optExact.IsChecked == true)
							{
								Exact(trans: true);
							}
							else if (wpfPasteTrans2.optLinks.IsChecked == true)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										Links(blnTranspose: true);
										goto end_IL_0078;
									}
									continue;
									end_IL_0078:
									break;
								}
							}
							else
							{
								NumberFormats(blnTranspose: true);
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							clsReporting.LogException(ex2);
							ProjectData.ClearProjectError();
						}
						clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(93562));
					}
					wpfPasteTrans2 = null;
					return;
				}
				}
			}
		}
		B();
	}

	private static void B()
	{
		Forms.WarningMessage(VH.A(93593));
	}

	private static void C()
	{
		Forms.WarningMessage(VH.A(93628));
	}
}
