using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.RowsColumns;

public sealed class Unhide
{
	public static void Rows(bool blnLogActivity = true)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		bool flag = default(bool);
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
			Application application = MH.A.Application;
			Range range = null;
			Range range2 = null;
			Range range3 = null;
			if (application.ScreenUpdating)
			{
				application.ScreenUpdating = false;
				flag = true;
			}
			Range range4;
			try
			{
				range3 = application.ActiveWindow.VisibleRange;
				range4 = (Range)((Range)application.Selection).Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
				if (Operators.ConditionalCompareObjectEqual(range4.Cells.CountLarge, 1, TextCompare: false))
				{
					if (!Conversions.ToBoolean(range4.Hidden))
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
						range = range4;
					}
				}
				else
				{
					range = A(range4);
				}
				range4.EntireRow.Hidden = false;
				if (range != null)
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
					range.EntireRow.Hidden = true;
				}
				if (Operators.ConditionalCompareObjectGreater(range4.Cells.CountLarge, 1, TextCompare: false))
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
					range2 = A(range4);
				}
				if (range != null)
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
					range.EntireRow.Hidden = false;
				}
				if (range2 != null)
				{
					{
						enumerator = range2.Areas.GetEnumerator();
						try
						{
							while (enumerator.MoveNext())
							{
								Range range5 = (Range)enumerator.Current;
								if (!Operators.ConditionalCompareObjectGreater(range5.EntireRow.OutlineLevel, 1, TextCompare: false))
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
								range5.Rows.ShowDetail = true;
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_01ba;
								}
								continue;
								end_IL_01ba:
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
				if (blnLogActivity)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						Core.LogActivity(VH.A(172006));
						break;
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			Window activeWindow = application.ActiveWindow;
			activeWindow.ScrollRow = Conversions.ToInteger(NewLateBinding.LateGet(range3.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41347), new object[0], null, null, null));
			activeWindow.ScrollColumn = Conversions.ToInteger(NewLateBinding.LateGet(range3.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41354), new object[0], null, null, null));
			_ = null;
			if (flag)
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
				application.ScreenUpdating = true;
			}
			range2 = null;
			range3 = null;
			range4 = null;
			range = null;
			application = null;
			return;
		}
	}

	public static void Columns(bool blnLogActivity = true)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Application application = MH.A.Application;
		Range range = null;
		Range range2 = null;
		Range range3 = null;
		bool flag = default(bool);
		if (application.ScreenUpdating)
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
			application.ScreenUpdating = false;
			flag = true;
		}
		Range range4;
		try
		{
			range3 = application.ActiveWindow.VisibleRange;
			range4 = (Range)((Range)application.Selection).Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
			if (Operators.ConditionalCompareObjectEqual(range4.Cells.CountLarge, 1, TextCompare: false))
			{
				if (!Conversions.ToBoolean(range4.Hidden))
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
					range = range4;
				}
			}
			else
			{
				range = A(range4);
			}
			range4.EntireColumn.Hidden = false;
			if (range != null)
			{
				range.EntireColumn.Hidden = true;
			}
			if (Operators.ConditionalCompareObjectGreater(range4.Cells.CountLarge, 1, TextCompare: false))
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
				range2 = A(range4);
			}
			if (range != null)
			{
				range.EntireColumn.Hidden = false;
			}
			if (range2 != null)
			{
				{
					IEnumerator enumerator = range2.Areas.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							Range range5 = (Range)enumerator.Current;
							if (!Operators.ConditionalCompareObjectGreater(range5.EntireColumn.OutlineLevel, 1, TextCompare: false))
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
							range5.Columns.ShowDetail = true;
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_01a8;
							}
							continue;
							end_IL_01a8:
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
			if (blnLogActivity)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					Core.LogActivity(VH.A(172029));
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		Window activeWindow = application.ActiveWindow;
		activeWindow.ScrollRow = Conversions.ToInteger(NewLateBinding.LateGet(range3.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41347), new object[0], null, null, null));
		activeWindow.ScrollColumn = Conversions.ToInteger(NewLateBinding.LateGet(range3.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41354), new object[0], null, null, null));
		_ = null;
		if (flag)
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
			application.ScreenUpdating = true;
		}
		range2 = null;
		range3 = null;
		range4 = null;
		range = null;
		application = null;
	}

	private static Range A(Range A)
	{
		Range result;
		try
		{
			result = A.SpecialCells(XlCellType.xlCellTypeVisible, RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = null;
			ProjectData.ClearProjectError();
		}
		return result;
	}
}
