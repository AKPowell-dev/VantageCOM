using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.ExcelApp;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.RowsColumns;

public sealed class Collapse
{
	public static void Rows()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Application application = MH.A.Application;
		if (EditMode.IsEditMode(application))
		{
			application = null;
			return;
		}
		Application application2 = application;
		try
		{
			if (application2.Windows.Count > 0)
			{
				IEnumerator enumerator = default(IEnumerator);
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
					if (application2.ActiveWindow.SelectedSheets.Count > 1)
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
						if (!Core.ConfirmMultipleSheets())
						{
							break;
						}
					}
					application2.ScreenUpdating = false;
					enumerator = application2.ActiveWindow.SelectedSheets.GetEnumerator();
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
								switch (7)
								{
								case 0:
									continue;
								}
								break;
							}
							((Worksheet)objectValue).Outline.ShowLevels(1, RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_00f2;
							}
							continue;
							end_IL_00f2:
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
					Core.LogActivity(VH.A(170154));
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
		finally
		{
			application2.ScreenUpdating = true;
		}
		application2 = null;
		application = null;
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
			switch (2)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Application application = MH.A.Application;
			if (EditMode.IsEditMode(application))
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						application = null;
						return;
					}
				}
			}
			Application application2 = application;
			try
			{
				if (application2.Windows.Count > 0)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						if (application2.ActiveWindow.SelectedSheets.Count > 1)
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
							if (!Core.ConfirmMultipleSheets())
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										goto end_IL_008e;
									}
									continue;
									end_IL_008e:
									break;
								}
								break;
							}
						}
						application2.ScreenUpdating = false;
						try
						{
							enumerator = application2.ActiveWindow.SelectedSheets.GetEnumerator();
							while (enumerator.MoveNext())
							{
								object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
								if (objectValue is Worksheet)
								{
									((Worksheet)objectValue).Outline.ShowLevels(RuntimeHelpers.GetObjectValue(Missing.Value), 1);
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
										continue;
									}
									(enumerator as IDisposable).Dispose();
									break;
								}
							}
						}
						Core.LogActivity(VH.A(170181));
						break;
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(ex2.Message);
				ProjectData.ClearProjectError();
			}
			finally
			{
				application2.ScreenUpdating = true;
			}
			application2 = null;
			application = null;
			return;
		}
	}
}
