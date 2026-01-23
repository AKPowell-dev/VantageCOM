using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.LogoLibrary.UI;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.LogoLibrary;

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
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (Access.IsLegacyPlan())
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								Access.NoLicenseMessage();
								return;
							}
						}
					}
					if (!Access.AllowPowerPointOperation((PlanType)6, (Restriction)0, false))
					{
						C();
					}
					else
					{
						A();
					}
					return;
				}
			}
		}
		B();
	}

	private static void A()
	{
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		//IL_002d: Expected O, but got Unknown
		Application application = NG.A.Application;
		DocumentWindow activeWindow = application.ActiveWindow;
		CustomTaskPane value = null;
		int hWND = activeWindow.HWND;
		Dpi = new clsDisplay();
		if (TaskPanes == null)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			TaskPanes = new Dictionary<int, CustomTaskPane>();
		}
		if (!TaskPanes.TryGetValue(hWND, out value))
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
			value = NG.A.CustomTaskPanes.Add(new ctpLogoLibrary(), Pane.PANE_TITLE, activeWindow);
			clsPanes.A(value, Dpi);
			value.VisibleChanged += A;
			TaskPanes.Add(hWND, value);
		}
		value.Visible = true;
		A((ctpLogoLibrary)value.Control).ShowPane();
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).RemoveEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).AddEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(A));
		activeWindow = null;
		value = null;
		application = null;
	}

	internal static void B()
	{
		if (NG.A.Application.Windows.Count <= 0)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			try
			{
				Pane.TaskPanes[NG.A.Application.ActiveWindow.HWND].Visible = false;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			Dictionary<int, CustomTaskPane> TaskPanes = Pane.TaskPanes;
			clsPanes.RemoveOrphanedPanes(ref TaskPanes, Pane.PANE_TITLE);
			Pane.TaskPanes = TaskPanes;
			return;
		}
	}

	private static void A(object A, EventArgs B)
	{
		CustomTaskPane customTaskPane = (CustomTaskPane)A;
		wpfLogoLibrary wpfLogoLibrary2 = Pane.A((ctpLogoLibrary)customTaskPane.Control);
		if (!customTaskPane.Visible)
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
			if (!KG.A)
			{
				wpfLogoLibrary2.HidePane();
				PB.Settings.TaskPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
				C();
			}
		}
		customTaskPane = null;
		wpfLogoLibrary2 = null;
	}

	private static wpfLogoLibrary A(ctpLogoLibrary A)
	{
		return (wpfLogoLibrary)A.ElementHost1.Child;
	}

	public static bool IsVisible()
	{
		return clsPanes.IsVisible(TaskPanes);
	}

	internal static void C()
	{
		KG.A.InvalidateControl(Pane.RIBBON_CONTROL);
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		Dictionary<int, CustomTaskPane> TaskPanes = Pane.TaskPanes;
		bool num = clsPanes.PresentationCloseFinal(ref TaskPanes, A);
		Pane.TaskPanes = TaskPanes;
		if (num)
		{
			C();
		}
	}
}
