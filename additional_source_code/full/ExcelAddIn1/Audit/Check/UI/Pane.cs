using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.UI;

public sealed class Pane
{
	private static readonly string m_A = VH.A(36112);

	private static readonly string m_B = VH.A(36149);

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
				switch (5)
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
							switch (2)
							{
							case 0:
								break;
							default:
								Access.NoLicenseMessage();
								return;
							}
						}
					}
					if (!Access.AllowExcelOperation((PlanType)5, (Restriction)0, false))
					{
						while (true)
						{
							switch (7)
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
			value.Visible = true;
		}
		else
		{
			value = MH.A.CustomTaskPanes.Add(new ctpAuditCheck(), Pane.m_A, activeWindow);
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
		}
		value = null;
		activeWindow = null;
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
					switch (6)
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
				switch (5)
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

	private static void A(object A, EventArgs B)
	{
		CustomTaskPane customTaskPane = (CustomTaskPane)A;
		if (!customTaskPane.Visible)
		{
			if (wpfAudit.InstanceRunningAnalysis != null)
			{
				wpfAudit.InstanceRunningAnalysis.A(A: true);
				string text;
				if (!KH.A)
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
					text = VH.A(35576);
				}
				else
				{
					text = VH.A(35826);
				}
				Forms.InfoMessage(text);
			}
			if (!KH.A)
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
				try
				{
					K.Settings.ExplorerPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
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
				int hwnd = application.ActiveWindow.Hwnd;
				if (PanesCollection != null)
				{
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

	internal static wpfAudit A()
	{
		wpfAudit result;
		try
		{
			if (PanesCollection == null)
			{
				result = null;
			}
			else
			{
				Application application = MH.A.Application;
				if (application.Windows.Count == 0)
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
						result = null;
						break;
					}
				}
				else
				{
					int hwnd = application.ActiveWindow.Hwnd;
					ctpAuditCheck obj = PanesCollection[hwnd].Control as ctpAuditCheck;
					object obj2;
					if (obj == null)
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
						obj2 = null;
					}
					else
					{
						obj2 = obj.A.AuditView;
					}
					result = (wpfAudit)obj2;
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result = null;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	internal static void B()
	{
		KH.A.InvalidateControl(Pane.m_B);
	}

	private static wpfPane A(ctpAuditCheck A)
	{
		return (wpfPane)A.ElementHost1.Child;
	}
}
