using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Threading;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes.Templated;

public sealed class Pane
{
	[CompilerGenerated]
	internal sealed class WD
	{
		public DispatcherTimer A;

		public ctpTemplatedShape A;

		[SpecialName]
		internal void A(object A, EventArgs B)
		{
			if (this.A == null)
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
				this.A.Stop();
				this.A = null;
				ctpTemplatedShape ctpTemplatedShape = this.A;
				ctpTemplatedShape.BringToFront();
				ctpTemplatedShape.Focus();
				_ = null;
				return;
			}
		}
	}

	public static readonly string PANE_TITLE = AH.A(70473);

	private static readonly string m_A = AH.A(70518);

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
		//IL_0059: Unknown result type (might be due to invalid IL or missing references)
		//IL_0063: Expected O, but got Unknown
		if (blnPressed && !Access.AllowPowerPointOperation((PlanType)6, (Restriction)0, false))
		{
			while (true)
			{
				switch (4)
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
		Application application = NG.A.Application;
		DocumentWindow activeWindow = application.ActiveWindow;
		CustomTaskPane value = null;
		int hWND = activeWindow.HWND;
		bool flag = false;
		if (blnPressed)
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
			Dpi = new clsDisplay();
		}
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
			if (TaskPanes.TryGetValue(hWND, out value))
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
				value.Visible = blnPressed;
				flag = true;
			}
		}
		if (blnPressed && !flag)
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
			ctpTemplatedShape control = new ctpTemplatedShape();
			value = NG.A.CustomTaskPanes.Add(control, PANE_TITLE, activeWindow);
			clsPanes.A(value, Dpi, 350);
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
			TaskPanes.Add(hWND, value);
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
			A((ctpTemplatedShape)null);
		}
		activeWindow = null;
		value = null;
		application = null;
	}

	private static void A(object A, EventArgs B)
	{
		CustomTaskPane customTaskPane = (CustomTaskPane)A;
		wpfTemplatedShape wpfTemplatedShape2 = Pane.A((ctpTemplatedShape)customTaskPane.Control);
		if (customTaskPane.Visible)
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
			wpfTemplatedShape2.ShowPane();
		}
		else if (!KG.A)
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
			wpfTemplatedShape2.HidePane();
			PB.Settings.TaskPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
			Pane.B();
		}
		customTaskPane = null;
		wpfTemplatedShape2 = null;
	}

	internal static void A()
	{
		if (NG.A.Application.Windows.Count <= 0)
		{
			return;
		}
		checked
		{
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
				int key = NG.A.Application.ActiveWindow.HWND;
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
							NG.A.CustomTaskPanes.Remove(TaskPanes[key]);
							TaskPanes.Remove(key);
							ProjectData.ClearProjectError();
						}
					}
					return;
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

	private static wpfTemplatedShape A(ctpTemplatedShape A)
	{
		return (wpfTemplatedShape)A.ElementHost1.Child;
	}

	public static bool IsVisible()
	{
		return clsPanes.IsVisible(TaskPanes);
	}

	internal static void B()
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
			switch (1)
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

	private static void A(ctpTemplatedShape A = null)
	{
		if (A == null)
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
			A = (ctpTemplatedShape)TaskPanes[NG.A.Application.ActiveWindow.HWND].Control;
		}
		if (A == null)
		{
			return;
		}
		DispatcherTimer A2;
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			A2 = new DispatcherTimer
			{
				Interval = new TimeSpan(1L)
			};
			A2.Tick += [SpecialName] (object obj, EventArgs B) =>
			{
				if (A2 != null)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
						{
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							A2.Stop();
							A2 = null;
							ctpTemplatedShape ctpTemplatedShape = A;
							ctpTemplatedShape.BringToFront();
							ctpTemplatedShape.Focus();
							_ = null;
							return;
						}
						}
					}
				}
			};
			A2.Start();
			return;
		}
	}
}
