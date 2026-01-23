using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class SumBar
{
	public static void Toggle(Range rng = null)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		if (rng == null)
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
			if (application.Selection is Range)
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
				rng = (Range)application.Selection;
			}
		}
		if (rng != null)
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
			if (!Base.IsWorksheetProtected(rng.Worksheet))
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
				if (Operators.ConditionalCompareObjectLessEqual(rng.Cells.CountLarge, Operators.MultiplyObject(10, ((Worksheet)application.ActiveSheet).Columns.CountLarge), TextCompare: false))
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
					application.ScreenUpdating = false;
					application.EnableEvents = false;
					try
					{
						bool num = JH.A(rng);
						if (SumBar.B(application.ActiveCell))
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
							A(rng);
						}
						else
						{
							B(rng);
						}
						if (num)
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
							JH.A(rng, VH.A(151417));
						}
						Base.LogActivity(VH.A(151417));
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						Base.HandleFormattingException(ex2);
						ProjectData.ClearProjectError();
					}
					application.ScreenUpdating = true;
					application.EnableEvents = true;
				}
				else
				{
					Forms.WarningMessage(VH.A(151432));
				}
			}
		}
		application = null;
	}

	private static void A(Range A)
	{
		Range range = A;
		range.ClearContents();
		range.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
		range.Font.ColorIndex = 1;
		if (((Range)range.Cells[1, 1]).Row > 1)
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
			Range range2 = ((Range)range.Cells[1, 1]).get_Offset((object)(-1), RuntimeHelpers.GetObjectValue(Missing.Value));
			range.VerticalAlignment = RuntimeHelpers.GetObjectValue(range2.VerticalAlignment);
			range.Font.Size = RuntimeHelpers.GetObjectValue(range2.Font.Size);
			range2 = null;
		}
		Range usedRange = range.Worksheet.UsedRange;
		int num = Conversions.ToInteger(Operators.SubtractObject(Operators.AddObject(usedRange.Column, usedRange.Columns.CountLarge), 1));
		usedRange = null;
		Range range3 = JH.A(A, (Microsoft.Office.Interop.Excel.Application)null);
		if (range3 != null)
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = range3.Rows.GetEnumerator();
				IEnumerator enumerator2 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					Range range4 = (Range)enumerator.Current;
					bool flag = false;
					try
					{
						enumerator2 = range4.EntireRow.Cells.GetEnumerator();
						while (true)
						{
							if (enumerator2.MoveNext())
							{
								Range range5 = (Range)enumerator2.Current;
								if (range5.Column > num)
								{
									break;
								}
								if (!SumBar.B(range5))
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
									flag = true;
									break;
								}
								break;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_01bf;
								}
								continue;
								end_IL_01bf:
								break;
							}
							break;
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								(enumerator2 as IDisposable).Dispose();
								break;
							}
						}
					}
					if (flag)
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
					range4.RowHeight = range.Worksheet.StandardHeight;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_0220;
					}
					continue;
					end_IL_0220:
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
			range3 = null;
		}
		range = null;
	}

	private static void B(Range A)
	{
		bool flag = false;
		if (SumBar.A(A))
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
			if (MessageBox.Show(VH.A(151537), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
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
				flag = true;
			}
		}
		if (flag)
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
			A.set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)VH.A(41385));
			A.Font.Underline = XlUnderlineStyle.xlUnderlineStyleSingleAccounting;
			A.Font.Color = KH.A.DefaultBorderColor;
			A.Font.Size = 11;
			A.VerticalAlignment = XlVAlign.xlVAlignBottom;
			A.RowHeight = 5;
			_ = null;
			return;
		}
	}

	private static bool A(Range A)
	{
		bool flag = false;
		if (Operators.ConditionalCompareObjectEqual(A.Cells.CountLarge, 1, TextCompare: false))
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
			if (Conversions.ToBoolean(A.HasFormula))
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
				flag = true;
			}
			else if (!string.IsNullOrEmpty(Conversions.ToString(A.Text)))
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
				flag = true;
			}
		}
		else
		{
			try
			{
				flag = A.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value)) != null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (!flag)
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
				try
				{
					flag = A.SpecialCells(XlCellType.xlCellTypeConstants, RuntimeHelpers.GetObjectValue(Missing.Value)) != null;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
			}
		}
		return flag;
	}

	internal static bool B(Range A)
	{
		return Conversions.ToBoolean(Conversions.ToBoolean(Operators.CompareObjectEqual(A.Font.Underline, XlUnderlineStyle.xlUnderlineStyleSingleAccounting, TextCompare: false)) && Conversions.ToBoolean(Operators.CompareObjectEqual(A.Value2, VH.A(41385), TextCompare: false)));
	}
}
