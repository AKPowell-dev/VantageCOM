using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.View;

public sealed class Splits
{
	public static void Split()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			if (application.Selection is Range)
			{
				try
				{
					Window activeWindow = application.ActiveWindow;
					if (activeWindow.SelectedSheets.Count <= 1)
					{
						goto IL_0095;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						break;
					}
					if (MessageBox.Show(VH.A(174246), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.Cancel)
					{
						goto IL_0095;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_0086;
						}
						continue;
						end_IL_0086:
						break;
					}
					goto end_IL_0040;
					IL_0095:
					application.ScreenUpdating = false;
					application.EnableEvents = false;
					try
					{
						string cell = ((Range)application.Selection).get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						string cell2 = application.ActiveCell.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						Worksheet worksheet = (Worksheet)application.ActiveSheet;
						Window window = activeWindow;
						window.Split = false;
						int scrollRow = window.ScrollRow;
						int scrollColumn = window.ScrollColumn;
						try
						{
							enumerator = window.SelectedSheets.GetEnumerator();
							while (enumerator.MoveNext())
							{
								object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
								if (!(objectValue is Worksheet))
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
								Worksheet obj = (Worksheet)objectValue;
								obj.Activate();
								window.Split = false;
								if (obj != worksheet)
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
									window.ScrollRow = scrollRow;
									window.ScrollColumn = scrollColumn;
								}
								((_Worksheet)obj).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value)).Select();
								((_Worksheet)obj).get_Range((object)cell2, RuntimeHelpers.GetObjectValue(Missing.Value)).Activate();
								window.Split = true;
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
						window = null;
						worksheet.Activate();
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					finally
					{
						Worksheet worksheet = null;
					}
					application.ScreenUpdating = true;
					application.EnableEvents = true;
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)7, VH.A(174386));
					end_IL_0040:;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				finally
				{
					Window activeWindow = null;
				}
			}
			application = null;
			return;
		}
	}

	public static void Unsplit()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			try
			{
				Window activeWindow = application.ActiveWindow;
				if (activeWindow.SelectedSheets.Count <= 1 || MessageBox.Show(VH.A(174419), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.Cancel)
				{
					application.ScreenUpdating = false;
					application.EnableEvents = false;
					try
					{
						object objectValue = RuntimeHelpers.GetObjectValue(application.ActiveSheet);
						enumerator = activeWindow.SelectedSheets.GetEnumerator();
						try
						{
							while (enumerator.MoveNext())
							{
								object objectValue2 = RuntimeHelpers.GetObjectValue(enumerator.Current);
								if (!(objectValue2 is Worksheet))
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
								NewLateBinding.LateCall(objectValue2, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
								application.ActiveWindow.Split = false;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_00fa;
								}
								continue;
								end_IL_00fa:
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
						NewLateBinding.LateCall(objectValue, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					finally
					{
						object objectValue = null;
					}
					application.ScreenUpdating = true;
					application.EnableEvents = true;
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)7, VH.A(174563));
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			finally
			{
				Window activeWindow = null;
			}
			application = null;
			return;
		}
	}
}
