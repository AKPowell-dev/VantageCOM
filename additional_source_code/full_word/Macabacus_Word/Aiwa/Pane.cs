using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Aiwa.UI;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Aiwa;

public sealed class Pane
{
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
					A();
					return;
				}
			}
		}
		B();
	}

	private static void A()
	{
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Expected O, but got Unknown
		Window activeWindow = PC.A.Application.ActiveWindow;
		CustomTaskPane value = null;
		int hwnd = activeWindow.Hwnd;
		Dpi = new clsDisplay();
		if (TaskPanes == null)
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
			TaskPanes = new Dictionary<int, CustomTaskPane>();
		}
		if (!TaskPanes.TryGetValue(hwnd, out value))
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
			value = PC.A.CustomTaskPanes.Add(new ctpAiwa(), Pane.PANE_TITLE, activeWindow);
			clsPanes.A(value, Dpi);
			value.VisibleChanged += A;
			TaskPanes.Add(hwnd, value);
		}
		value.Visible = true;
		A((ctpAiwa)value.Control).ShowPane();
		activeWindow = null;
		value = null;
	}

	internal static void B()
	{
		if (PC.A.Application.Windows.Count <= 0)
		{
			return;
		}
		checked
		{
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
				int key = PC.A.Application.ActiveWindow.Hwnd;
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
							PC.A.CustomTaskPanes.Remove(TaskPanes[key]);
							TaskPanes.Remove(key);
							ProjectData.ClearProjectError();
						}
					}
					while (true)
					{
						switch (2)
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
		wpfHome wpfHome2 = Pane.A((ctpAiwa)customTaskPane.Control);
		if (!customTaskPane.Visible)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (!NC.B)
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
				wpfHome2.HidePane();
				N.Settings.TaskPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
				C();
			}
		}
		customTaskPane = null;
		wpfHome2 = null;
	}

	private static wpfHome A(ctpAiwa A)
	{
		return (wpfHome)A.ElementHost1.Child;
	}

	public static bool IsVisible()
	{
		return clsPanes.IsVisible(TaskPanes);
	}

	internal static void C()
	{
		NC.A.InvalidateControl(Pane.RIBBON_CONTROL);
	}
}
