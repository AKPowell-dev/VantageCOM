using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Audit.Visualizations;
using ExcelAddIn1.ExcelApp;
using ExcelAddIn1.Format;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class Flatten
{
	[CompilerGenerated]
	private static string m_A;

	private static string LastFunctionFlattened
	{
		[CompilerGenerated]
		get
		{
			return Flatten.m_A;
		}
		[CompilerGenerated]
		set
		{
			Flatten.m_A = value;
		}
	} = "";

	public static void Selection()
	{
		if (!A())
		{
			return;
		}
		Range range = null;
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		if (EditMode.IsEditMode(application))
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					application = null;
					return;
				}
			}
		}
		Range activeCell;
		try
		{
			activeCell = application.ActiveCell;
			range = (Range)application.Selection;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.WarningMessage(VH.A(156927));
			application = null;
			ProjectData.ClearProjectError();
			return;
		}
		try
		{
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			bool flag = JH.A(range);
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = range.Areas.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range2 = (Range)enumerator.Current;
					range2.Value2 = RuntimeHelpers.GetObjectValue(range2.Value2);
					foreach (Range item in range2)
					{
						AutoColor.AutoColorIfNotEmpty(item);
					}
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_0112;
					}
					continue;
					end_IL_0112:
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
			range.Select();
			activeCell.Activate();
			if (flag)
			{
				JH.A(range, VH.A(497));
			}
			Common.RefreshLiveVisualizations(range);
			A(VH.A(157030));
		}
		catch (OutOfMemoryException ex3)
		{
			ProjectData.SetProjectError(ex3);
			OutOfMemoryException ex4 = ex3;
			Forms.ErrorMessage(ex4.Message);
			ProjectData.ClearProjectError();
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			clsReporting.LogException(ex6);
			Forms.ErrorMessage(ex6.Message);
			ProjectData.ClearProjectError();
		}
		finally
		{
			application.EnableEvents = true;
			application.ScreenUpdating = true;
			application = null;
			range = null;
			activeCell = null;
		}
	}

	public static void FlattenFunction()
	{
		if (!A())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
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
			Common.ClearVisualizations();
			wpfFlattenFunction wpfFlattenFunction2 = new wpfFlattenFunction();
			if (LastFunctionFlattened.Length > 0)
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
				wpfFlattenFunction2.txtFunction.Text = LastFunctionFlattened;
				wpfFlattenFunction2.txtFunction.SelectAll();
			}
			wpfFlattenFunction2.ShowDialog();
			if (wpfFlattenFunction2.DialogResult.HasValue)
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
				if (wpfFlattenFunction2.DialogResult.Value)
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
					int C = 0;
					string text = wpfFlattenFunction2.txtFunction.Text;
					Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
					if (EditMode.IsEditMode(application))
					{
						application = null;
						return;
					}
					Microsoft.Office.Interop.Excel.Application application2 = application;
					XlCalculation calculation = application2.Calculation;
					application2.Calculation = XlCalculation.xlCalculationManual;
					application2.ScreenUpdating = false;
					application2.EnableEvents = false;
					_ = null;
					try
					{
						if (wpfFlattenFunction2.radSelection.IsChecked == true)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								if (!(application.Selection is Range))
								{
									break;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									Range range = Helpers.SpecialCellsFormulas((Range)application.Selection);
									if (range == null)
									{
										break;
									}
									while (true)
									{
										switch (6)
										{
										case 0:
											continue;
										}
										A(range, text, ref C);
										range = null;
										break;
									}
									break;
								}
								break;
							}
						}
						else if (wpfFlattenFunction2.radSheets.IsChecked == true)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								enumerator = application.ActiveWindow.SelectedSheets.GetEnumerator();
								try
								{
									while (enumerator.MoveNext())
									{
										object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
										if (objectValue is Worksheet)
										{
											A((Worksheet)objectValue, text, ref C);
										}
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											break;
										default:
											goto end_IL_01eb;
										}
										continue;
										end_IL_01eb:
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
								break;
							}
						}
						else
						{
							try
							{
								enumerator2 = application.ActiveWorkbook.Worksheets.GetEnumerator();
								while (enumerator2.MoveNext())
								{
									A((Worksheet)enumerator2.Current, text, ref C);
								}
								while (true)
								{
									switch (6)
									{
									case 0:
										break;
									default:
										goto end_IL_0249;
									}
									continue;
									end_IL_0249:
									break;
								}
							}
							finally
							{
								if (enumerator2 is IDisposable)
								{
									while (true)
									{
										switch (4)
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
					}
					catch (OutOfMemoryException ex)
					{
						ProjectData.SetProjectError(ex);
						OutOfMemoryException ex2 = ex;
						Forms.ErrorMessage(ex2.Message);
						ProjectData.ClearProjectError();
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						clsReporting.LogException(ex4);
						Forms.ErrorMessage(ex4.Message);
						ProjectData.ClearProjectError();
					}
					LastFunctionFlattened = text;
					Forms.SuccessMessage(VH.A(157065) + C + VH.A(157084));
					A(VH.A(157129));
					Microsoft.Office.Interop.Excel.Application application3 = application;
					application3.Calculation = calculation;
					application3.ScreenUpdating = true;
					application3.EnableEvents = true;
					_ = null;
					application = null;
				}
			}
			wpfFlattenFunction2 = null;
			return;
		}
	}

	private static void A(Worksheet A, string B, ref int C)
	{
		Range range = null;
		if (!A.ProtectContents)
		{
			try
			{
				range = A.UsedRange.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (range == null)
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
				Flatten.A(range, B, ref C);
				return;
			}
		}
		Forms.WarningMessage(VH.A(157162) + A.Name + VH.A(157249));
	}

	private static void A(Range A, string B, ref int C)
	{
		checked
		{
			foreach (Range item in A)
			{
				if (Conversions.ToBoolean(NewLateBinding.LateGet(item, null, VH.A(46494), new object[0], null, null, null)))
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
					if (Helpers.IsFunctionMatch(item, B))
					{
						long num = Conversions.ToLong(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(item, null, VH.A(103802), new object[0], null, null, null), null, VH.A(62391), new object[0], null, null, null), null, VH.A(152052), new object[0], null, null, null));
						NewLateBinding.LateSetComplex(NewLateBinding.LateGet(item, null, VH.A(103802), new object[0], null, null, null), null, VH.A(57302), new object[1] { NewLateBinding.LateGet(NewLateBinding.LateGet(item, null, VH.A(103802), new object[0], null, null, null), null, VH.A(57302), new object[0], null, null, null) }, null, null, OptimisticSet: false, RValueBase: true);
						C = (int)(C + num);
					}
				}
				else
				{
					if (!Helpers.IsFunctionMatch(item, B))
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
					item.Value2 = RuntimeHelpers.GetObjectValue(item.Value2);
					C++;
				}
			}
		}
	}

	public static void IsolateSheets()
	{
		if (!A())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Range range = null;
		Range range2 = null;
		int num = 0;
		Common.ClearVisualizations(application);
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook;
		try
		{
			activeWorkbook = application.ActiveWorkbook;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			application = null;
			ProjectData.ClearProjectError();
			return;
		}
		Microsoft.Office.Interop.Excel.Sheets selectedSheets = application.ActiveWindow.SelectedSheets;
		if (!activeWorkbook.Saved)
		{
			string text;
			if (selectedSheets.Count == 1)
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
				text = VH.A(157302);
			}
			else
			{
				text = VH.A(157347) + selectedSheets.Count + VH.A(157356);
			}
			if (MessageBox.Show(VH.A(157397) + text + VH.A(157464), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						selectedSheets = null;
						activeWorkbook = null;
						application = null;
						return;
					}
				}
			}
		}
		object objectValue = RuntimeHelpers.GetObjectValue(activeWorkbook.ActiveSheet);
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
			range = (Range)application.Selection;
			range2 = application.ActiveCell;
		}
		XlDisplayDrawingObjects displayDrawingObjects = activeWorkbook.DisplayDrawingObjects;
		Microsoft.Office.Interop.Excel.Application application2 = application;
		XlCalculation calculation = application2.Calculation;
		application2.Calculation = XlCalculation.xlCalculationManual;
		application2.ScreenUpdating = false;
		application2.EnableEvents = false;
		_ = null;
		activeWorkbook.DisplayDrawingObjects = XlDisplayDrawingObjects.xlHide;
		checked
		{
			Worksheet worksheet;
			try
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = selectedSheets.GetEnumerator();
					IEnumerator enumerator2 = default(IEnumerator);
					IEnumerator enumerator3 = default(IEnumerator);
					while (enumerator.MoveNext())
					{
						object objectValue2 = RuntimeHelpers.GetObjectValue(enumerator.Current);
						if (!(objectValue2 is Worksheet))
						{
							continue;
						}
						worksheet = (Worksheet)objectValue2;
						if (worksheet.ProtectContents)
						{
							Forms.WarningMessage(VH.A(157529) + worksheet.Name + VH.A(157249));
							continue;
						}
						Range range3 = null;
						try
						{
							range3 = worksheet.UsedRange.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
						}
						if (range3 == null)
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
						try
						{
							enumerator2 = range3.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Range range4 = (Range)enumerator2.Current;
								if (!range4.Formula.ToString().Contains(VH.A(157572)))
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
									if (!A(range4))
									{
										continue;
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											continue;
										}
										break;
									}
								}
								if (Conversions.ToBoolean(NewLateBinding.LateGet(range4, null, VH.A(46494), new object[0], null, null, null)))
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
									Range range5 = (Range)NewLateBinding.LateGet(range4, null, VH.A(103802), new object[0], null, null, null);
									List<Tuple<Range, object>> list = new List<Tuple<Range, object>>();
									try
									{
										enumerator3 = range5.Cells.GetEnumerator();
										while (enumerator3.MoveNext())
										{
											object objectValue3 = RuntimeHelpers.GetObjectValue(enumerator3.Current);
											list.Add(new Tuple<Range, object>((Range)objectValue3, RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue3, null, VH.A(57302), new object[0], null, null, null))));
										}
										while (true)
										{
											switch (2)
											{
											case 0:
												break;
											default:
												goto end_IL_035e;
											}
											continue;
											end_IL_035e:
											break;
										}
									}
									finally
									{
										if (enumerator3 is IDisposable)
										{
											while (true)
											{
												switch (4)
												{
												case 0:
													continue;
												}
												(enumerator3 as IDisposable).Dispose();
												break;
											}
										}
									}
									using List<Tuple<Range, object>>.Enumerator enumerator4 = list.GetEnumerator();
									while (enumerator4.MoveNext())
									{
										Tuple<Range, object> current = enumerator4.Current;
										current.Item1.Value2 = RuntimeHelpers.GetObjectValue(current.Item2);
										num++;
									}
									while (true)
									{
										switch (6)
										{
										case 0:
											break;
										default:
											goto end_IL_03ce;
										}
										continue;
										end_IL_03ce:
										break;
									}
								}
								else
								{
									range4.Value2 = RuntimeHelpers.GetObjectValue(range4.Value2);
									num++;
								}
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_0411;
								}
								continue;
								end_IL_0411:
								break;
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									(enumerator2 as IDisposable).Dispose();
									break;
								}
							}
						}
						range3 = null;
						activeWorkbook.Activate();
						worksheet.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
						activeWorkbook.DisplayDrawingObjects = XlDisplayDrawingObjects.xlDisplayShapes;
						worksheet.ClearArrows();
						activeWorkbook.DisplayDrawingObjects = XlDisplayDrawingObjects.xlHide;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_047c;
						}
						continue;
						end_IL_047c:
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
			catch (OutOfMemoryException ex5)
			{
				ProjectData.SetProjectError(ex5);
				OutOfMemoryException ex6 = ex5;
				Forms.ErrorMessage(ex6.Message);
				ProjectData.ClearProjectError();
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				clsReporting.LogException(ex8);
				Forms.ErrorMessage(ex8.Message);
				ProjectData.ClearProjectError();
			}
			activeWorkbook.DisplayDrawingObjects = displayDrawingObjects;
			activeWorkbook.Activate();
			NewLateBinding.LateCall(objectValue, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
			if (range != null)
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
				range2.Activate();
			}
			Microsoft.Office.Interop.Excel.Application application3 = application;
			application3.Calculation = calculation;
			application3.ScreenUpdating = true;
			application3.EnableEvents = true;
			_ = null;
			Forms.SuccessMessage(VH.A(157065) + num + VH.A(157084));
			A(VH.A(157585));
			range = null;
			range2 = null;
			worksheet = null;
			activeWorkbook = null;
			selectedSheets = null;
			application = null;
		}
	}

	private static bool A(Range A)
	{
		Worksheet worksheet = A.Worksheet;
		bool flag = false;
		checked
		{
			Range range;
			try
			{
				A.ShowPrecedents(RuntimeHelpers.GetObjectValue(Missing.Value));
				int num = 1;
				int num2 = 1;
				bool flag2 = true;
				while (true)
				{
					range = null;
					try
					{
						range = (Range)A.NavigateArrow(true, num, num2);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
						goto IL_00f1;
					}
					if (Operators.CompareString(A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) != 0)
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
						if (range.Worksheet == worksheet)
						{
							flag2 = false;
							num2++;
							continue;
						}
						flag = true;
					}
					goto IL_00f1;
					IL_00f1:
					if (flag)
					{
						break;
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
					if (!flag2)
					{
						num2 = 1;
						flag2 = true;
						num++;
						continue;
					}
					break;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			range = null;
			worksheet = null;
			return flag;
		}
	}

	private static bool A()
	{
		return Access.AllowExcelOperation((PlanType)4, (Restriction)1, false);
	}

	private static void A(string A)
	{
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, A);
	}
}
