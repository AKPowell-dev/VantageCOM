using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Threading;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.UI;

public sealed class Pane
{
	[CompilerGenerated]
	internal sealed class SF
	{
		public DispatcherTimer A;

		public ctpSuperFind2 A;

		[SpecialName]
		internal void A(object A, EventArgs B)
		{
			if (this.A != null)
			{
				this.A.Stop();
				this.A = null;
				ctpSuperFind2 ctpSuperFind = this.A;
				ctpSuperFind.BringToFront();
				ctpSuperFind.Focus();
				_ = null;
			}
		}
	}

	private static readonly string m_A = VH.A(140471);

	private static readonly string m_B = VH.A(140492);

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

	internal static Dictionary<int, CustomTaskPane> PanesCollection
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
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (!Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								B();
								return;
							}
						}
					}
					A();
					return;
				}
			}
		}
		Hide();
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
			switch (5)
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
			if (PanesCollection == null)
			{
				PanesCollection = new Dictionary<int, CustomTaskPane>();
			}
			if (PanesCollection.TryGetValue(hwnd, out var value))
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
				value.Visible = true;
			}
			else
			{
				value = MH.A.CustomTaskPanes.Add(new ctpSuperFind2(), Pane.m_A, activeWindow);
				CustomTaskPane customTaskPane = value;
				try
				{
					customTaskPane.DockPosition = (MsoCTPDockPosition)K.Settings.ExplorerPanePosn;
					clsPanes.A(value, Dpi, K.Settings.ExplorerPaneWidth);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				customTaskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;
				customTaskPane.Visible = true;
				customTaskPane.VisibleChanged += A;
				customTaskPane.DockPositionChanged += B;
				customTaskPane = null;
				PanesCollection.Add(hwnd, value);
				A((ctpSuperFind2)null);
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
		int key = MH.A.Application.ActiveWindow.Hwnd;
		try
		{
			PanesCollection[key].Visible = false;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		checked
		{
			try
			{
				for (int i = PanesCollection.Count - 1; i >= 0; i += -1)
				{
					try
					{
						key = PanesCollection.ElementAt(i).Key;
						_ = PanesCollection[key].Title;
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						MH.A.CustomTaskPanes.Remove(PanesCollection[key]);
						PanesCollection.Remove(key);
						ProjectData.ClearProjectError();
					}
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
					break;
				}
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			if (PanesCollection == null)
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
				if (PanesCollection.Any())
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
					PanesCollection = null;
					return;
				}
			}
		}
	}

	private static void A(ctpSuperFind2 A = null)
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
			A = (ctpSuperFind2)PanesCollection[MH.A.Application.ActiveWindow.Hwnd].Control;
		}
		if (A == null)
		{
			return;
		}
		DispatcherTimer A2;
		while (true)
		{
			switch (6)
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
					A2.Stop();
					A2 = null;
					ctpSuperFind2 ctpSuperFind = A;
					ctpSuperFind.BringToFront();
					ctpSuperFind.Focus();
					_ = null;
				}
			};
			A2.Start();
			return;
		}
	}

	private static void A(object A, EventArgs B)
	{
		CustomTaskPane customTaskPane = (CustomTaskPane)A;
		if (customTaskPane.Visible)
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
			Pane.A(customTaskPane.Control as ctpSuperFind2);
			try
			{
				wpfResults wpfResults2 = Pane.A((ctpSuperFind2)customTaskPane.Control);
				if (wpfResults2.SearchResults.Count > 0)
				{
					wpfResults2.A(MH.A.Application);
				}
				wpfResults2 = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		else
		{
			try
			{
				Pane.A((ctpSuperFind2)customTaskPane.Control).Q();
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			if (!KH.A)
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
					Pane.A((ctpSuperFind2)customTaskPane.Control).I();
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
				try
				{
					K.Settings.ExplorerPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
				}
				catch (Exception ex7)
				{
					ProjectData.SetProjectError(ex7);
					Exception ex8 = ex7;
					ProjectData.ClearProjectError();
				}
				Pane.B();
			}
		}
		customTaskPane = null;
	}

	private static void B(object A, EventArgs B)
	{
		CustomTaskPane customTaskPane = (CustomTaskPane)A;
		K.Settings.ExplorerPanePosn = (int)customTaskPane.DockPosition;
		try
		{
			Pane.A((ctpSuperFind2)customTaskPane.Control).DockPosition = customTaskPane.DockPosition;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		customTaskPane = null;
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
				int hwnd = application.ActiveWindow.Hwnd;
				if (PanesCollection != null)
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
					try
					{
						return PanesCollection[hwnd].Visible;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						MH.A.CustomTaskPanes.Remove(PanesCollection[hwnd]);
						PanesCollection.Remove(hwnd);
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
		KH.A.InvalidateControl(Pane.m_B);
	}

	private static wpfPane A(ctpSuperFind2 A)
	{
		return (wpfPane)A.ElementHost1.Child;
	}

	private static wpfResults A(ctpSuperFind2 A)
	{
		return Pane.A(A).ResultsView;
	}
}
