using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Threading;
using A;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts.GrowthArrow;

public sealed class Pane
{
	[CompilerGenerated]
	internal sealed class ED
	{
		public DispatcherTimer A;

		public ctpGrowthArrow A;

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
				ctpGrowthArrow ctpGrowthArrow = this.A;
				ctpGrowthArrow.BringToFront();
				ctpGrowthArrow.Focus();
				_ = null;
				return;
			}
		}
	}

	private static readonly string m_A = VH.A(52976);

	private static readonly string m_B = VH.A(53001);

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
			if (!Licensing.AllowChartAddOnOperation())
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
						C();
						return;
					}
				}
			}
			A();
		}
		else
		{
			B();
		}
	}

	private static void A()
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_000a: Expected O, but got Unknown
		Dpi = new clsDisplay();
		if (MH.A.Application.Windows.Count > 0)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Window activeWindow = MH.A.Application.ActiveWindow;
					int hwnd = activeWindow.Hwnd;
					if (TaskPanes == null)
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
						TaskPanes = new Dictionary<int, CustomTaskPane>();
					}
					if (!TaskPanes.TryGetValue(hwnd, out var value))
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
						value = MH.A.CustomTaskPanes.Add(new ctpGrowthArrow(), Pane.m_A, activeWindow);
						clsPanes.A(value, Dpi, 320);
						value.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
						value.VisibleChanged += A;
						TaskPanes.Add(hwnd, value);
					}
					A((ctpGrowthArrow)value.Control).ShowPane();
					value.Visible = true;
					A((ctpGrowthArrow)null);
					value = null;
					activeWindow = null;
					return;
				}
				}
			}
		}
		C();
	}

	internal static void B()
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
						switch (1)
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
		wpfPane wpfPane2 = Pane.A((ctpGrowthArrow)customTaskPane.Control);
		if (!customTaskPane.Visible)
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
			wpfPane2.HidePane();
			K.Settings.TaskPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
			C();
		}
		customTaskPane = null;
	}

	private static wpfPane A(ctpGrowthArrow A)
	{
		return (wpfPane)A.ElementHost1.Child;
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
				int hwnd = application.ActiveWindow.Hwnd;
				if (TaskPanes != null)
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

	internal static void C()
	{
		KH.A.InvalidateControl(Pane.m_B);
	}

	private static void A(ctpGrowthArrow A = null)
	{
		if (A == null)
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
			A = (ctpGrowthArrow)TaskPanes[MH.A.Application.ActiveWindow.Hwnd].Control;
		}
		if (A == null)
		{
			return;
		}
		DispatcherTimer A2;
		while (true)
		{
			switch (4)
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
							ctpGrowthArrow ctpGrowthArrow = A;
							ctpGrowthArrow.BringToFront();
							ctpGrowthArrow.Focus();
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
