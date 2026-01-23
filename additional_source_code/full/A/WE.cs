using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using ExcelAddIn1.Formulas;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

internal sealed class WE
{
	private Application m_A;

	private Worksheet m_A;

	private Worksheet m_B;

	internal WE(Application A)
	{
		this.m_A = A;
		this.m_A = null;
		this.m_B = null;
	}

	internal void A(Range A, Range B, Range C = null, Workbook D = null)
	{
		checked
		{
			try
			{
				Worksheet worksheet = B.Worksheet;
				Worksheet worksheet2 = A.Worksheet;
				if (this.m_A == null)
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
					worksheet2.Copy(worksheet, RuntimeHelpers.GetObjectValue(Missing.Value));
					this.m_A = (Worksheet)this.m_A.ActiveSheet;
					WE.A(this.m_A);
				}
				else
				{
					this.m_A.UsedRange.Clear();
					Range usedRange = worksheet2.UsedRange;
					usedRange.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
					((Range)this.m_A.Cells[usedRange.Row, usedRange.Column]).PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				Range rows = A.Rows;
				Range columns = A.Columns;
				int row = A.Row;
				int column = A.Column;
				int count = rows.Count;
				int count2 = columns.Count;
				Range cell = (Range)this.m_A.Cells[row, column];
				Range cell2 = (Range)this.m_A.Cells[row - 1 + count, column - 1 + count2];
				Range range = ((_Worksheet)this.m_A).get_Range((object)cell, (object)cell2);
				if (this.m_B == null)
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
					if (D == null)
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
						D = (Workbook)worksheet.Parent;
					}
					Sheets sheets = D.Sheets;
					this.m_B = (Worksheet)sheets.Add(worksheet, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					WE.A(this.m_B);
				}
				else
				{
					this.m_B.UsedRange.Clear();
				}
				row = B.Row;
				column = B.Column;
				Range destination = (Range)this.m_B.Cells[row, column];
				range.Cut(destination);
				cell = (Range)this.m_B.Cells[row, column];
				cell2 = (Range)this.m_B.Cells[row - 1 + count, column - 1 + count2];
				((_Worksheet)this.m_B).get_Range((object)cell, (object)cell2).Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
				B.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				if (C == null)
				{
					int C2 = 0;
					int D2 = 0;
					C = JH.A(B, A, out C2, out D2);
				}
				this.A(C);
			}
			finally
			{
				Range destination = null;
				Sheets sheets = null;
				Range cell2 = null;
				Range cell = null;
				Range rows = null;
				Range usedRange = null;
				Worksheet worksheet = null;
				Worksheet worksheet2 = null;
				C = null;
				D = null;
			}
		}
	}

	private static void A(Worksheet A)
	{
		int num = 0;
		while (true)
		{
			num = checked(num + 1);
			try
			{
				A.Name = string.Format(VH.A(87072), num);
				break;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				if (num >= 1000)
				{
					Forms.WarningMessage(string.Format(VH.A(87097), ex2.Message));
					throw;
				}
				ProjectData.ClearProjectError();
			}
		}
	}

	private void A(Range A)
	{
		string text = string.Format(VH.A(87293), this.m_A.Name.Replace(VH.A(39851), VH.A(39854)));
		try
		{
			Range range = Helpers.SpecialCellsFormulas(A);
			if (range == null)
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
						return;
					}
				}
			}
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = range.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range2 = (Range)enumerator.Current;
					try
					{
						if (Operators.CompareString(range2.PrefixCharacter.ToString(), VH.A(39851), TextCompare: false) == 0)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_009b;
								}
								continue;
								end_IL_009b:
								break;
							}
							continue;
						}
						string text2 = Helpers.B(range2);
						if (!text2.Contains(text))
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
							Helpers.A(range2, text2.Replace(text, ""));
							break;
						}
					}
					finally
					{
						range2 = null;
					}
				}
				while (true)
				{
					switch (3)
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
		finally
		{
			Range range = null;
		}
	}

	internal void A()
	{
		B(this.m_A);
		this.m_A = null;
		B(this.m_B);
		this.m_B = null;
		this.m_A = null;
	}

	private void B(Worksheet A)
	{
		if (A == null)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			string name;
			try
			{
				name = A.Name;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
				return;
			}
			bool displayAlerts = this.m_A.DisplayAlerts;
			try
			{
				this.m_A.DisplayAlerts = false;
				A.Delete();
				return;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Forms.ErrorMessage(string.Format(VH.A(87306), name));
				clsReporting.LogException(ex4);
				ProjectData.ClearProjectError();
				return;
			}
			finally
			{
				this.m_A.DisplayAlerts = displayAlerts;
			}
		}
	}
}
