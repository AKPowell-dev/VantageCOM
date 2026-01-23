using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Timers;
using A;
using ExcelAddIn1.Model;
using MacabacusMacros.ExcelHelpers;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.UndoRedo;

public sealed class Core
{
	[CompilerGenerated]
	internal sealed class MG
	{
		public string A;

		public MG(MG A)
		{
			if (A == null)
			{
				return;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A(object A, ElapsedEventArgs B)
		{
			try
			{
				Core.A(this.A);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			finally
			{
				Core.m_A.Dispose();
				Core.m_A = null;
			}
		}
	}

	private static CellProp m_A;

	private static Dictionary<Microsoft.Office.Interop.Excel.Workbook, WorkbookItem> m_A;

	private static Timer m_A;

	public static void Enable()
	{
		Core.m_A = new Dictionary<Microsoft.Office.Interop.Excel.Workbook, WorkbookItem>(new WorkbookComparer());
	}

	public static void Disable()
	{
		try
		{
			IEnumerator enumerator = MH.A.Application.Workbooks.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					A((Microsoft.Office.Interop.Excel.Workbook)enumerator.Current);
				}
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
			Core.m_A = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void SaveToUndoStack(Range ChangedRange, string UndoAction)
	{
		Range range = JH.A(ChangedRange);
		int num = Conversions.ToInteger(range.Cells.CountLarge);
		if (num > KH.A.UndoMaxCells)
		{
			return;
		}
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
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
				SavedCell[] array = new SavedCell[num + 1];
				int num2 = 0;
				Worksheet worksheet = ChangedRange.Worksheet;
				_ = worksheet.Application;
				Microsoft.Office.Interop.Excel.Workbook key = (Microsoft.Office.Interop.Excel.Workbook)worksheet.Parent;
				if (Core.m_A == null)
				{
					Core.m_A = new Dictionary<Microsoft.Office.Interop.Excel.Workbook, WorkbookItem>(new WorkbookComparer());
				}
				WorkbookItem workbookItem = Core.m_A[key];
				workbookItem.RedoStack.Clear();
				enumerator = range.GetEnumerator();
				Dictionary<Range, CellProp> value;
				CellProp value2;
				SavedCell savedCell;
				try
				{
					while (enumerator.MoveNext())
					{
						Range range2 = (Range)enumerator.Current;
						savedCell = new SavedCell(range2);
						try
						{
							if (workbookItem.BaseSheets.TryGetValue(worksheet, out value))
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									if (value.TryGetValue(range2, out value2))
									{
										try
										{
											savedCell.OldProp.A(ref value2);
											CellProp cellProp = value2;
											SavedCell savedCell2;
											CellProp A = (savedCell2 = savedCell).NewProp;
											cellProp.A(ref A);
											savedCell2.NewProp = A;
										}
										catch (Exception ex)
										{
											ProjectData.SetProjectError(ex);
											Exception ex2 = ex;
											ProjectData.ClearProjectError();
										}
										break;
									}
									try
									{
										value2 = new CellProp();
										savedCell.OldProp.A(ref value2);
										CellProp cellProp2 = value2;
										SavedCell savedCell2;
										CellProp A = (savedCell2 = savedCell).NewProp;
										cellProp2.A(ref A);
										savedCell2.NewProp = A;
										value.Add(range2, value2);
									}
									catch (Exception ex3)
									{
										ProjectData.SetProjectError(ex3);
										Exception ex4 = ex3;
										ProjectData.ClearProjectError();
									}
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
						array[num2] = savedCell;
						num2++;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_01b8;
						}
						continue;
						end_IL_01b8:
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
				StackItem obj = new StackItem(array, ChangedRange, UndoAction);
				workbookItem.UndoStack.Push(obj);
				num = workbookItem.UndoStack.Count;
				if (num > 100)
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
						Stack source = new Stack(workbookItem.UndoStack);
						workbookItem.UndoStack.Clear();
						int num3 = num - 1;
						for (num2 = 1; num2 <= num3; num2++)
						{
							workbookItem.UndoStack.Push(RuntimeHelpers.GetObjectValue(source.Cast<object>().ElementAtOrDefault(num2)));
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							source = null;
							break;
						}
					}
					catch (Exception ex7)
					{
						ProjectData.SetProjectError(ex7);
						Exception ex8 = ex7;
						ProjectData.ClearProjectError();
					}
				}
				Core.A(UndoAction);
				range = null;
				key = null;
				workbookItem = null;
				value = null;
				value2 = null;
				savedCell = null;
				return;
			}
		}
	}

	private static void A(Microsoft.Office.Interop.Excel.Workbook A)
	{
		try
		{
			WorkbookItem workbookItem = Core.m_A[A];
			workbookItem.UndoStack.Clear();
			workbookItem.RedoStack.Clear();
			workbookItem.BaseSheets.Clear();
			_ = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void B(Microsoft.Office.Interop.Excel.Workbook A)
	{
		checked
		{
			for (int i = Core.m_A.Keys.Count - 1; i >= 0; i += -1)
			{
				try
				{
					_ = Core.m_A.Keys.ElementAtOrDefault(i).Name;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					try
					{
						Core.m_A.Remove(Core.m_A.Keys.ElementAtOrDefault(i));
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						Interaction.MsgBox(VH.A(173809));
						Interaction.MsgBox(ex4.Message);
						ProjectData.ClearProjectError();
					}
					ProjectData.ClearProjectError();
				}
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
				return;
			}
		}
	}

	public static bool IndexSelection(Range Target)
	{
		if (KH.A.UndoEnabled)
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
			Range range = JH.A(Target);
			IEnumerator enumerator = default(IEnumerator);
			bool result = default(bool);
			if (!Operators.ConditionalCompareObjectGreater(range.Cells.CountLarge, KH.A.UndoMaxCells, TextCompare: false))
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
					{
						Worksheet worksheet = Target.Worksheet;
						Microsoft.Office.Interop.Excel.Workbook key = (Microsoft.Office.Interop.Excel.Workbook)worksheet.Parent;
						Dictionary<Range, CellProp> value = new Dictionary<Range, CellProp>(new RangeComparer());
						WorkbookItem value2;
						try
						{
							if (Core.m_A.TryGetValue(key, out value2))
							{
								if (value2.BaseSheets.TryGetValue(worksheet, out value))
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
									try
									{
										enumerator = range.GetEnumerator();
										while (enumerator.MoveNext())
										{
											Range A = (Range)enumerator.Current;
											try
											{
												CellProp cellProp = new CellProp();
												cellProp.A(ref A);
												value.Remove(A);
												value.Add(A, cellProp);
												cellProp = null;
											}
											catch (Exception ex)
											{
												ProjectData.SetProjectError(ex);
												Exception ex2 = ex;
												ProjectData.ClearProjectError();
											}
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
													break;
												default:
													(enumerator as IDisposable).Dispose();
													goto end_IL_0129;
												}
												continue;
												end_IL_0129:
												break;
											}
										}
									}
								}
								else
								{
									try
									{
										value2.BaseSheets.Add(worksheet, Core.A(range));
									}
									catch (Exception ex3)
									{
										ProjectData.SetProjectError(ex3);
										Exception ex4 = ex3;
										ProjectData.ClearProjectError();
									}
								}
							}
							else
							{
								try
								{
									value2 = new WorkbookItem();
									value2.BaseSheets.Add(worksheet, Core.A(range));
									Core.m_A.Add(key, value2);
								}
								catch (Exception ex5)
								{
									ProjectData.SetProjectError(ex5);
									Exception ex6 = ex5;
									ProjectData.ClearProjectError();
								}
							}
							result = true;
						}
						catch (Exception ex7)
						{
							ProjectData.SetProjectError(ex7);
							Exception ex8 = ex7;
							ProjectData.ClearProjectError();
						}
						range = null;
						worksheet = null;
						key = null;
						value = null;
						value2 = null;
						return result;
					}
					}
				}
			}
		}
		bool result2 = default(bool);
		return result2;
	}

	private static Dictionary<Range, CellProp> A(Range A)
	{
		Dictionary<Range, CellProp> dictionary = new Dictionary<Range, CellProp>(new RangeComparer());
		IEnumerator enumerator = default(IEnumerator);
		CellProp cellProp;
		try
		{
			enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range A2 = (Range)enumerator.Current;
				cellProp = new CellProp();
				cellProp.A(ref A2);
				dictionary.Add(A2, cellProp);
			}
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
		cellProp = null;
		return dictionary;
	}

	public static void Undo()
	{
		MG a = default(MG);
		MG CS_0024_003C_003E8__locals2 = new MG(a);
		Application application = MH.A.Application;
		XlCalculation calculation = application.Calculation;
		bool flag = false;
		bool flag2 = false;
		checked
		{
			WorkbookItem workbookItem = default(WorkbookItem);
			StackItem stackItem = default(StackItem);
			try
			{
				Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application.ActiveWorkbook;
				workbookItem = Core.m_A[activeWorkbook];
				if (workbookItem.UndoStack.Count <= 0)
				{
					return;
				}
				application.EnableEvents = false;
				application.ScreenUpdating = false;
				flag2 = application.CutCopyMode == XlCutCopyMode.xlCopy;
				application.Calculation = XlCalculation.xlCalculationManual;
				flag = true;
				stackItem = (StackItem)workbookItem.UndoStack.Pop();
				workbookItem.RedoStack.Push(stackItem);
				StackItem stackItem2 = stackItem;
				try
				{
					stackItem2.Range.Worksheet.Activate();
					_ = workbookItem.BaseSheets[stackItem2.Range.Worksheet];
					try
					{
						int num = Information.UBound(stackItem2.Cells) - 1;
						for (int i = 0; i <= num; i++)
						{
							SavedCell savedCell = stackItem2.Cells[i];
							CellProp oldProp = savedCell.OldProp;
							SavedCell savedCell2;
							Range A = (savedCell2 = savedCell).Range;
							oldProp.B(ref A);
							savedCell2.Range = A;
							savedCell = null;
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
							Core.A(stackItem2.Range);
							break;
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					workbookItem.RedoStack.Pop();
					Undo();
					ProjectData.ClearProjectError();
				}
				B(stackItem2.Name);
				stackItem2 = null;
			}
			catch (KeyNotFoundException ex5)
			{
				ProjectData.SetProjectError(ex5);
				KeyNotFoundException ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
			finally
			{
				try
				{
					if (workbookItem.UndoStack.Count > 0)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							stackItem = (StackItem)workbookItem.UndoStack.Peek();
							CS_0024_003C_003E8__locals2.A = stackItem.Name;
							Core.m_A = new Timer(1.0);
							Core.m_A.Elapsed += [SpecialName] (object obj, ElapsedEventArgs B) =>
							{
								try
								{
									Core.A(CS_0024_003C_003E8__locals2.A);
								}
								catch (Exception ex11)
								{
									ProjectData.SetProjectError(ex11);
									Exception ex12 = ex11;
									ProjectData.ClearProjectError();
								}
								finally
								{
									Core.m_A.Dispose();
									Core.m_A = null;
								}
							};
							Core.m_A.AutoReset = false;
							Core.m_A.Enabled = true;
							break;
						}
					}
				}
				catch (Exception ex9)
				{
					ProjectData.SetProjectError(ex9);
					Exception ex10 = ex9;
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
					application.EnableEvents = true;
					application.ScreenUpdating = true;
					application.Calculation = calculation;
					B(stackItem.Name);
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
						Paste.CopiedRange.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
					}
				}
				application = null;
				Microsoft.Office.Interop.Excel.Workbook activeWorkbook = null;
				workbookItem = null;
				stackItem = null;
			}
		}
	}

	public static void Redo()
	{
		Application application = MH.A.Application;
		XlCalculation calculation = application.Calculation;
		bool flag = false;
		bool flag2 = false;
		checked
		{
			WorkbookItem workbookItem = default(WorkbookItem);
			try
			{
				Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application.ActiveWorkbook;
				workbookItem = Core.m_A[activeWorkbook];
				if (workbookItem.RedoStack.Count <= 0)
				{
					return;
				}
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
					application.EnableEvents = false;
					application.ScreenUpdating = false;
					flag2 = application.CutCopyMode == XlCutCopyMode.xlCopy;
					application.Calculation = XlCalculation.xlCalculationManual;
					flag = true;
					StackItem stackItem = (StackItem)workbookItem.RedoStack.Pop();
					workbookItem.UndoStack.Push(stackItem);
					StackItem stackItem2 = stackItem;
					try
					{
						stackItem2.Range.Worksheet.Activate();
						_ = workbookItem.BaseSheets[stackItem2.Range.Worksheet];
						try
						{
							int num = Information.UBound(stackItem2.Cells) - 1;
							for (int i = 0; i <= num; i++)
							{
								SavedCell savedCell = stackItem2.Cells[i];
								CellProp newProp = savedCell.NewProp;
								SavedCell savedCell2;
								Range A = (savedCell2 = savedCell).Range;
								newProp.B(ref A);
								savedCell2.Range = A;
								savedCell = null;
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								Core.A(stackItem2.Range);
								break;
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						workbookItem.UndoStack.Pop();
						Redo();
						ProjectData.ClearProjectError();
					}
					Core.A(stackItem2.Name);
					stackItem2 = null;
					return;
				}
			}
			catch (KeyNotFoundException ex5)
			{
				ProjectData.SetProjectError(ex5);
				KeyNotFoundException ex6 = ex5;
				Core.A();
				ProjectData.ClearProjectError();
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
			finally
			{
				StackItem stackItem;
				try
				{
					if (workbookItem.RedoStack.Count > 0)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							stackItem = (StackItem)workbookItem.RedoStack.Peek();
							B(stackItem.Name);
							break;
						}
					}
				}
				catch (Exception ex9)
				{
					ProjectData.SetProjectError(ex9);
					Exception ex10 = ex9;
					ProjectData.ClearProjectError();
				}
				if (flag)
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
					application.EnableEvents = true;
					application.ScreenUpdating = true;
					application.Calculation = calculation;
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
						Paste.CopiedRange.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
					}
				}
				application = null;
				Microsoft.Office.Interop.Excel.Workbook activeWorkbook = null;
				workbookItem = null;
				stackItem = null;
			}
		}
	}

	private static void A(string A)
	{
		MH.A.Application.OnUndo(A, clsUtilities.XLAM_FILE_NAME + VH.A(173842));
	}

	private static void B(string A)
	{
		MH.A.Application.OnRepeat(A, clsUtilities.XLAM_FILE_NAME + VH.A(173853));
	}

	public static int UndoStackSize()
	{
		return Core.m_A[MH.A.Application.ActiveWorkbook].UndoStack.Count;
	}

	private static void A(Range A)
	{
		Ranges.ScrollIntoView(A);
		A.Select();
	}

	private static void A()
	{
		Disable();
		Enable();
	}

	public static bool UndoButtonEnabled()
	{
		return Core.m_A[MH.A.Application.ActiveWorkbook].UndoStack.Count > 0;
	}

	public static bool RedoButtonEnabled()
	{
		return Core.m_A[MH.A.Application.ActiveWorkbook].RedoStack.Count > 0;
	}
}
