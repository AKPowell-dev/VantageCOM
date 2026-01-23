using System;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Xml;
using A;
using ExcelAddIn1.Audit;
using ExcelAddIn1.Audit.TraceDialogs;
using ExcelAddIn1.Keyboard;
using ExcelAddIn1.UndoRedo;
using ExcelAddIn1.View;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

public sealed class clsUtilities
{
	public static readonly string XLAM_FILE_NAME = VH.A(198529);

	public static readonly string ADDIN_NAME = VH.A(40448);

	public static void StartupProcedures(XmlDocument xmlSettings)
	{
		Application application = MH.A.Application;
		KH.A = new clsSettings(xmlSettings);
		try
		{
			Shortcuts.BuildDictionary();
			Shortcuts.Load();
			if (application.UserControl)
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
				if (Workspace.MaximizeOnStartUp(xmlSettings))
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
					Workspace.Maximize(blnRequireAuthentication: false);
				}
			}
			if (K.Settings.AutoTracePrecedents)
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
				new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).AddEventHandler(application, new AppEvents_SheetSelectionChangeEventHandler(AutoTrace.AutoTracePrecedentsEventHandler));
			}
			else if (K.Settings.AutoTraceDependents)
			{
				new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).AddEventHandler(application, new AppEvents_SheetSelectionChangeEventHandler(AutoTrace.AutoTraceDependentsEventHandler));
			}
			if (KH.A.UndoEnabled)
			{
				ExcelAddIn1.UndoRedo.Core.Enable();
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		Base.WorkshareLoaded = Base.IsWorkshareLoaded(application);
		application = null;
	}

	public static string MacabacusXlamPath()
	{
		return Path.Combine(clsEnvironment.CommonAppDataPath, XLAM_FILE_NAME);
	}

	public static void InitializeMacabacus()
	{
		Application application = MH.A.Application;
		AddIn addIn = null;
		string text = MacabacusXlamPath();
		bool flag = false;
		try
		{
			addIn = application.AddIns.get_Item((object)ADDIN_NAME);
			if (Operators.CompareString(addIn.FullName.ToLower(), text.ToLower(), TextCompare: false) != 0)
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
					Forms.ErrorMessage(VH.A(198066) + XLAM_FILE_NAME + VH.A(198075) + text);
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Microsoft.Office.Interop.Excel.Workbook workbook = default(Microsoft.Office.Interop.Excel.Workbook);
			foreach (Microsoft.Office.Interop.Excel.Workbook workbook2 in application.Workbooks)
			{
				workbook = workbook2;
				if (!workbook.Windows[1].Visible)
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
			if (!flag)
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
				workbook = application.Workbooks.Add(RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			try
			{
				addIn = application.AddIns.Add(text, RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Forms.ErrorMessage(VH.A(198333) + XLAM_FILE_NAME + VH.A(198364) + text + VH.A(198377) + ex4.Message);
				ProjectData.ClearProjectError();
			}
			if (!flag)
			{
				workbook.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				workbook = null;
			}
			ProjectData.ClearProjectError();
		}
		try
		{
			if (addIn != null)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					if (addIn.Installed)
					{
						break;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						addIn.Installed = true;
						break;
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
		addIn = null;
		application = null;
	}
}
