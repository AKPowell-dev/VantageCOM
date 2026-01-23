using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using A;
using ExcelAddIn1.Model;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.View;

public sealed class Gridlines
{
	public static void Toggle()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		while (true)
		{
			switch (3)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			bool flag = false;
			Application application = MH.A.Application;
			if (application.ActiveSheet is Worksheet)
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
					flag = application.CutCopyMode == XlCutCopyMode.xlCopy;
					Window activeWindow = application.ActiveWindow;
					activeWindow.DisplayGridlines = !activeWindow.DisplayGridlines;
					_ = null;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					clsReporting.LogException(ex2);
					ProjectData.ClearProjectError();
				}
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
					if (Paste.CopiedRange != null)
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
						Paste.CopiedRange.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
					}
				}
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)7, VH.A(175062));
			}
			application = null;
			return;
		}
	}

	public static void AutoHide(bool blnEnabled)
	{
		Application application = MH.A.Application;
		if (blnEnabled)
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
			new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(125628)).AddEventHandler(application, new AppEvents_SheetActivateEventHandler(A));
			A(RuntimeHelpers.GetObjectValue(application.ActiveSheet));
		}
		else
		{
			new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(125628)).RemoveEventHandler(application, new AppEvents_SheetActivateEventHandler(A));
		}
		application = null;
	}

	private static void A(object A)
	{
		if (!(A is Worksheet))
		{
			return;
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
			try
			{
				Application application = ((Worksheet)A).Application;
				if (application.ActiveWindow.DisplayGridlines)
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
					if (application.ActiveWindow.Visible)
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
						if (application.CutCopyMode == (XlCutCopyMode)0)
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
							application.ActiveWindow.DisplayGridlines = false;
						}
					}
				}
				application = null;
				return;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
				return;
			}
		}
	}
}
