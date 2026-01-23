using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;

namespace ExcelAddIn1.Library2.Versioning;

public sealed class Pane
{
	private static readonly string m_A = VH.A(83090);

	private static readonly string m_B = VH.A(83121);

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
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Check.B();
					B();
					return;
				}
			}
		}
		CustomTaskPane value = null;
		if (TaskPanes == null)
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
			if (!TaskPanes.TryGetValue(MH.A.Application.ActiveWindow.Hwnd, out value))
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
				value.Visible = false;
				return;
			}
		}
	}

	internal static void A()
	{
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Expected O, but got Unknown
		Window activeWindow = MH.A.Application.ActiveWindow;
		int hwnd = activeWindow.Hwnd;
		CustomTaskPane value = null;
		bool flag = false;
		Dpi = new clsDisplay();
		if (TaskPanes != null)
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
			if (TaskPanes.TryGetValue(hwnd, out value))
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
				if (!value.Visible)
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
					value.Visible = true;
				}
				else
				{
					A((ctpVersions)value.Control).A();
				}
				flag = true;
			}
		}
		if (!flag)
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
			value = MH.A.CustomTaskPanes.Add(new ctpVersions(), Pane.m_A, activeWindow);
			clsPanes.A(value, Dpi);
			value.VisibleChanged += A;
			value.Visible = true;
			if (TaskPanes == null)
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
				TaskPanes = new Dictionary<int, CustomTaskPane>();
			}
			TaskPanes.Add(hwnd, value);
		}
		activeWindow = null;
		value = null;
	}

	private static void A(object A, EventArgs B)
	{
		CustomTaskPane customTaskPane = (CustomTaskPane)A;
		wpfVersions wpfVersions2 = Pane.A((ctpVersions)customTaskPane.Control);
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
			wpfVersions2.A();
		}
		else if (!KH.A)
		{
			wpfVersions2.B();
			K.Settings.TaskPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
			Pane.B();
		}
		customTaskPane = null;
		wpfVersions2 = null;
	}

	private static wpfVersions A(ctpVersions A)
	{
		return (wpfVersions)A.ElementHost1.Child;
	}

	public static bool IsVisible()
	{
		return clsPanes.A(TaskPanes);
	}

	internal static void B()
	{
		KH.A.InvalidateControl(Pane.m_B);
	}
}
