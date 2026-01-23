using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Publishing.Share;

public sealed class Pane
{
	public static readonly string PANE_TITLE = VH.A(98770);

	private static readonly string m_A = VH.A(98803);

	[CompilerGenerated]
	private static clsDisplay m_A;

	[CompilerGenerated]
	private static Dictionary<int, CustomTaskPane> m_A;

	private static clsDisplay Dpi
	{
		[CompilerGenerated]
		get
		{
			return Pane.m_A;
		}
		[CompilerGenerated]
		set
		{
			Pane.m_A = value;
		}
	}

	private static Dictionary<int, CustomTaskPane> TaskPanes
	{
		[CompilerGenerated]
		get
		{
			return Pane.m_A;
		}
		[CompilerGenerated]
		set
		{
			Pane.m_A = value;
		}
	} = null;

	public static void Toggle(bool blnPressed)
	{
		if (blnPressed)
		{
			if (!Access.AllowExcelOperation((PlanType)5, (Restriction)2, false))
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						B();
						return;
					}
				}
			}
			A();
		}
		else
		{
			Hide();
		}
	}

	private static void A()
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_000a: Expected O, but got Unknown
		Dpi = new clsDisplay();
		if (MH.A.Application.Windows.Count <= 0)
		{
			return;
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
			Window activeWindow = MH.A.Application.ActiveWindow;
			int hwnd = activeWindow.Hwnd;
			if (TaskPanes == null)
			{
				TaskPanes = new Dictionary<int, CustomTaskPane>();
			}
			if (TaskPanes.TryGetValue(hwnd, out var value))
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
				value.Visible = true;
			}
			else
			{
				value = MH.A.CustomTaskPanes.Add(new ctpShare(), PANE_TITLE, activeWindow);
				clsPanes.A(value, Dpi);
				CustomTaskPane customTaskPane = value;
				customTaskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
				customTaskPane.VisibleChanged += A;
				customTaskPane.Visible = true;
				_ = null;
				TaskPanes.Add(hwnd, value);
			}
			value = null;
			activeWindow = null;
			return;
		}
	}

	public static void Hide()
	{
		if (MH.A.Application.Windows.Count <= 0)
		{
			return;
		}
		checked
		{
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
				int key = MH.A.Application.ActiveWindow.Hwnd;
				try
				{
					TaskPanes[key].Visible = false;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				try
				{
					for (int i = TaskPanes.Count - 1; i >= 0; i += -1)
					{
						try
						{
							key = TaskPanes.ElementAt(i).Key;
							_ = TaskPanes[key].Title;
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							MH.A.CustomTaskPanes.Remove(TaskPanes[key]);
							TaskPanes.Remove(key);
							ProjectData.ClearProjectError();
						}
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
					return;
				}
			}
		}
	}

	private static void A(object A, EventArgs B)
	{
		CustomTaskPane customTaskPane = (CustomTaskPane)A;
		wpfShare wpfShare2 = Pane.A((ctpShare)customTaskPane.Control);
		if (customTaskPane.Visible)
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
			wpfShare2.ShowPane();
		}
		else
		{
			wpfShare2.HidePane();
			K.Settings.TaskPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
			Pane.B();
		}
		customTaskPane = null;
	}

	private static wpfShare A(ctpShare A)
	{
		return (wpfShare)A.ElementHost1.Child;
	}

	public static bool IsVisible()
	{
		try
		{
			Application application = MH.A.Application;
			if (application.Windows.Count > 0)
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
				int hwnd = application.ActiveWindow.Hwnd;
				if (TaskPanes != null)
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
					try
					{
						return TaskPanes[hwnd].Visible;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						MH.A.CustomTaskPanes.Remove(TaskPanes[hwnd]);
						TaskPanes.Remove(hwnd);
						ProjectData.ClearProjectError();
					}
				}
			}
			application = null;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		return false;
	}

	private static void B()
	{
		KH.A.InvalidateControl(Pane.m_A);
	}
}
