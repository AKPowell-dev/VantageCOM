using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;

namespace PowerPointAddIn1.Library2.Versioning;

public sealed class Pane
{
	private static readonly string m_A = AH.A(59027);

	private static readonly string m_B = AH.A(59058);

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
		Application application = NG.A.Application;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).RemoveEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(A));
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Check.B();
			B();
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).AddEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(A));
		}
		else
		{
			CustomTaskPane value = null;
			if (TaskPanes != null)
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
				if (TaskPanes.TryGetValue(application.ActiveWindow.HWND, out value))
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
					value.Visible = false;
				}
			}
		}
		application = null;
	}

	internal static void A()
	{
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		//IL_002b: Expected O, but got Unknown
		DocumentWindow activeWindow = NG.A.Application.ActiveWindow;
		int hWND = activeWindow.HWND;
		CustomTaskPane value = null;
		bool flag = false;
		Dpi = new clsDisplay();
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (TaskPanes.TryGetValue(hWND, out value))
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
				if (!value.Visible)
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
			value = NG.A.CustomTaskPanes.Add(new ctpVersions(), Pane.m_A, activeWindow);
			clsPanes.A(value, Dpi);
			value.VisibleChanged += A;
			value.Visible = true;
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
				TaskPanes = new Dictionary<int, CustomTaskPane>();
			}
			TaskPanes.Add(hWND, value);
		}
		_ = NG.A.Application;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).RemoveEventHandler(NG.A.Application, new EApplication_PresentationCloseFinalEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).AddEventHandler(NG.A.Application, new EApplication_PresentationCloseFinalEventHandler(A));
		_ = null;
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
			wpfVersions2.A();
		}
		else if (!KG.A)
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
			wpfVersions2.B();
			PB.Settings.TaskPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
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
		return clsPanes.IsVisible(TaskPanes);
	}

	internal static void B()
	{
		KG.A.InvalidateControl(Pane.m_B);
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
			switch (3)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			B();
			return;
		}
	}
}
