using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.RowsColumns;

public sealed class Group
{
	public static void Rows()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Application application = MH.A.Application;
			object obj = null;
			application.ScreenUpdating = false;
			try
			{
				if (application.ActiveWindow.SelectedSheets.Count > 1 && !Core.ConfirmMultipleSheets())
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_0059;
						}
						continue;
						end_IL_0059:
						break;
					}
				}
				else
				{
					Range range = (Range)application.Selection;
					if (Operators.ConditionalCompareObjectEqual(range.Rows.CountLarge, range.Worksheet.Rows.CountLarge, TextCompare: false))
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							B();
							break;
						}
					}
					else
					{
						string cell = range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)false, RuntimeHelpers.GetObjectValue(Missing.Value));
						range = null;
						obj = RuntimeHelpers.GetObjectValue(application.ActiveSheet);
						try
						{
							enumerator = application.ActiveWindow.SelectedSheets.GetEnumerator();
							while (enumerator.MoveNext())
							{
								object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
								if (objectValue is Worksheet)
								{
									Worksheet obj2 = (Worksheet)objectValue;
									obj2.Activate();
									((_Worksheet)obj2).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value)).Rows.Group(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									_ = null;
								}
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_0188;
								}
								continue;
								end_IL_0188:
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
						Core.LogActivity(VH.A(171358));
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (obj != null)
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
				NewLateBinding.LateCall(obj, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
			}
			application.ScreenUpdating = true;
			application = null;
			obj = null;
			return;
		}
	}

	private static void A()
	{
		Range range = default(Range);
		try
		{
			range = (Range)MH.A.Application.Selection;
			if (Operators.ConditionalCompareObjectEqual(range.Rows.CountLarge, range.Worksheet.Rows.CountLarge, TextCompare: false))
			{
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
					B();
					break;
				}
			}
			else
			{
				range.Rows.Group(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		JH.A((object)range);
	}

	private static void B()
	{
		Forms.WarningMessage(VH.A(171379));
	}

	public static void Columns()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Application application = MH.A.Application;
			object obj = null;
			application.ScreenUpdating = false;
			try
			{
				if (application.ActiveWindow.SelectedSheets.Count <= 1)
				{
					goto IL_006c;
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
				if (Core.ConfirmMultipleSheets())
				{
					goto IL_006c;
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_005d;
					}
					continue;
					end_IL_005d:
					break;
				}
				goto end_IL_0035;
				IL_006c:
				Range range = (Range)application.Selection;
				if (Operators.ConditionalCompareObjectEqual(range.Columns.CountLarge, range.Worksheet.Columns.CountLarge, TextCompare: false))
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						D();
						break;
					}
				}
				else
				{
					string cell = range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)false, RuntimeHelpers.GetObjectValue(Missing.Value));
					range = null;
					obj = RuntimeHelpers.GetObjectValue(application.ActiveSheet);
					{
						enumerator = application.ActiveWindow.SelectedSheets.GetEnumerator();
						try
						{
							while (enumerator.MoveNext())
							{
								object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
								if (!(objectValue is Worksheet))
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
								Worksheet obj2 = (Worksheet)objectValue;
								obj2.Activate();
								((_Worksheet)obj2).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value)).Columns.Group(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								_ = null;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_0198;
								}
								continue;
								end_IL_0198:
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
					Core.LogActivity(VH.A(171430));
				}
				end_IL_0035:;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (obj != null)
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
				NewLateBinding.LateCall(obj, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
			}
			application.ScreenUpdating = true;
			application = null;
			obj = null;
			return;
		}
	}

	private static void C()
	{
		Range range = default(Range);
		try
		{
			range = (Range)MH.A.Application.Selection;
			if (Operators.ConditionalCompareObjectEqual(range.Columns.CountLarge, range.Worksheet.Columns.CountLarge, TextCompare: false))
			{
				D();
			}
			else
			{
				range.Columns.Group(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		JH.A((object)range);
	}

	private static void D()
	{
		Forms.WarningMessage(VH.A(171457));
	}
}
