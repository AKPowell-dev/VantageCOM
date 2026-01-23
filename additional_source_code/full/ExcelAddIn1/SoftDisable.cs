using System;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Audit;
using ExcelAddIn1.Audit.Check.UI;
using ExcelAddIn1.Audit.Visualizations;
using ExcelAddIn1.Charts.GrowthArrow;
using ExcelAddIn1.ExcelApp;
using ExcelAddIn1.Formulas;
using ExcelAddIn1.Keyboard;
using ExcelAddIn1.Library2.UI;
using ExcelAddIn1.Model;
using ExcelAddIn1.Publishing.Share;
using ExcelAddIn1.RowsColumns;
using ExcelAddIn1.SuperFind2.UI;
using ExcelAddIn1.UndoRedo;
using ExcelAddIn1.View;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

public sealed class SoftDisable
{
	[CompilerGenerated]
	private static bool m_A;

	[CompilerGenerated]
	private static bool m_B;

	[CompilerGenerated]
	private static bool C;

	[CompilerGenerated]
	private static bool D;

	[CompilerGenerated]
	private static bool E;

	public static bool IsEnabled
	{
		[CompilerGenerated]
		get
		{
			return SoftDisable.m_A;
		}
		[CompilerGenerated]
		set
		{
			SoftDisable.m_A = value;
		}
	} = true;

	private static bool NavAidEnabled
	{
		[CompilerGenerated]
		get
		{
			return SoftDisable.m_B;
		}
		[CompilerGenerated]
		set
		{
			SoftDisable.m_B = value;
		}
	}

	private static bool WorkspaceMaximized
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	private static bool AutoTracePrecedents
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[CompilerGenerated]
		set
		{
			D = value;
		}
	}

	private static bool AutoTraceDependents
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[CompilerGenerated]
		set
		{
			E = value;
		}
	}

	public static void SoftToggle()
	{
		if (IsEnabled)
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
			if (!Access.AllowExcelOperation((PlanType)4, (Restriction)0, false) || Forms.OkCancelMessage2(VH.A(141529)) == DialogResult.Cancel)
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
				break;
			}
		}
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		application.ScreenUpdating = false;
		if (IsEnabled)
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
				A(application);
				application.EnableEvents = true;
				Environment.DisplayAlertsOn(application);
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)14, VH.A(141612));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			IsEnabled = false;
		}
		else
		{
			try
			{
				B(application);
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				clsReporting.LogException(ex4);
				ProjectData.ClearProjectError();
			}
			IsEnabled = true;
		}
		KH.A.Invalidate();
		application.ScreenUpdating = true;
		application = null;
	}

	public static Bitmap GetButtonImage()
	{
		if (IsEnabled)
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
					return J.RecordingPause;
				}
			}
		}
		return J.RecordingPlay;
	}

	public static string GetButtonLabel()
	{
		if (IsEnabled)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return VH.A(141643);
				}
			}
		}
		return VH.A(141654);
	}

	private static void A(Microsoft.Office.Interop.Excel.Application A)
	{
		Events.Remove();
		Shortcuts.Remove();
		ExcelAddIn1.UndoRedo.Core.Disable();
		NavAidEnabled = NavAid.Enabled;
		if (NavAidEnabled)
		{
			NavAid.Disable();
		}
		AutoTracePrecedents = K.Settings.AutoTracePrecedents;
		if (AutoTracePrecedents)
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
			AutoTrace.TogglePrecedents();
		}
		AutoTraceDependents = K.Settings.AutoTraceDependents;
		if (AutoTraceDependents)
		{
			AutoTrace.ToggleDependents();
		}
		ExcelAddIn1.View.Gridlines.AutoHide(blnEnabled: false);
		if (Translate.IsTranslating())
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
			Translate.ToggleTranslate();
		}
		Common.ClearVisualizations();
		WorkspaceMaximized = Workspace.Maximized;
		if (WorkspaceMaximized)
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
			Workspace.Maximize(blnRequireAuthentication: false);
		}
		RemoveTaskPanes();
		JH.A((object)Paste.CopiedRange);
		JH.A((object)Paste.CutRange);
		JH.A((object)Info.CopiedRange);
		Paste.CopiedRange = null;
		Paste.CutRange = null;
		Info.CopiedRange = null;
	}

	private static void B(Microsoft.Office.Interop.Excel.Application A)
	{
		Events.Add();
		Shortcuts.Load();
		if (KH.A.UndoEnabled)
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
			ExcelAddIn1.UndoRedo.Core.Enable();
		}
		if (NavAidEnabled)
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
			NavAid.Toggle(blnChecked: true);
		}
		if (AutoTracePrecedents)
		{
			AutoTrace.TogglePrecedents();
		}
		else if (AutoTraceDependents)
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
			AutoTrace.ToggleDependents();
		}
		if (Conversions.ToBoolean(KH.A.SettingsXml.DocumentElement.SelectSingleNode(VH.A(141667)).InnerText))
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
			ExcelAddIn1.View.Gridlines.AutoHide(blnEnabled: true);
		}
		if (!WorkspaceMaximized)
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
			Workspace.Maximize(blnRequireAuthentication: false);
			return;
		}
	}

	public static void RemoveTaskPanes()
	{
		ExcelAddIn1.Library2.UI.Pane.SoftDisable();
		ShortcutManager.Toggle(blnPressed: false);
		ExcelAddIn1.SuperFind2.UI.Pane.Hide();
		ExcelAddIn1.Audit.Check.UI.Pane.Hide();
		ExcelAddIn1.Publishing.Share.Pane.Hide();
		ExcelAddIn1.Charts.GrowthArrow.Pane.B();
		clsDiscuss.DiscussPaneToggle(blnPressed: false);
		CustomTaskPaneCollection customTaskPanes = MH.A.CustomTaskPanes;
		checked
		{
			for (int i = customTaskPanes.Count - 1; i >= 0; i += -1)
			{
				CustomTaskPane customTaskPane = customTaskPanes[i];
				customTaskPanes.Remove(customTaskPane);
				customTaskPane.Dispose();
				JH.A((object)customTaskPane);
				customTaskPane = null;
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
				customTaskPanes = null;
				return;
			}
		}
	}
}
