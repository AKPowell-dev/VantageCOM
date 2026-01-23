using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;

namespace PowerPointAddIn1.Colors.Recolor;

public sealed class Pane
{
	public static readonly string PANE_TITLE = AH.A(12720);

	private static readonly string m_A = AH.A(12735);

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
		//IL_005b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Expected O, but got Unknown
		if (blnPressed && !Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
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
					A();
					return;
				}
			}
		}
		Application application = NG.A.Application;
		DocumentWindow activeWindow = application.ActiveWindow;
		CustomTaskPane value = null;
		int hWND = activeWindow.HWND;
		bool flag = false;
		if (blnPressed)
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
			Dpi = new clsDisplay();
		}
		if (TaskPanes != null && TaskPanes.TryGetValue(hWND, out value))
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
			value.Visible = blnPressed;
			flag = true;
		}
		if (blnPressed)
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
			if (!flag)
			{
				ctpRecolor control = new ctpRecolor();
				value = NG.A.CustomTaskPanes.Add(control, PANE_TITLE, activeWindow);
				clsPanes.A(value, Dpi, 350);
				value.VisibleChanged += A;
				value.Visible = true;
				if (TaskPanes == null)
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
					TaskPanes = new Dictionary<int, CustomTaskPane>();
				}
				TaskPanes.Add(hWND, value);
			}
		}
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).RemoveEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(A));
		if (blnPressed)
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
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).AddEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(A));
		}
		activeWindow = null;
		value = null;
		application = null;
	}

	private static void A(object A, EventArgs B)
	{
		CustomTaskPane customTaskPane = (CustomTaskPane)A;
		wpfRecolor wpfRecolor2 = Pane.A((ctpRecolor)customTaskPane.Control);
		if (customTaskPane.Visible)
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
			wpfRecolor2.ShowPane();
		}
		else if (!KG.A)
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
			wpfRecolor2.HidePane();
			PB.Settings.TaskPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
			Pane.A();
		}
		customTaskPane = null;
		wpfRecolor2 = null;
	}

	private static wpfRecolor A(ctpRecolor A)
	{
		return (wpfRecolor)A.ElementHost1.Child;
	}

	public static bool IsVisible()
	{
		return clsPanes.IsVisible(TaskPanes);
	}

	private static void A()
	{
		KG.A.InvalidateControl(Pane.m_A);
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		Dictionary<int, CustomTaskPane> TaskPanes = Pane.TaskPanes;
		bool num = clsPanes.PresentationCloseFinal(ref TaskPanes, A);
		Pane.TaskPanes = TaskPanes;
		if (!num)
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
			Pane.A();
			return;
		}
	}
}
