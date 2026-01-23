using System;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Keyboard;

public sealed class ShortcutManager
{
	private static readonly string m_A = VH.A(161338);

	private static readonly string m_B = VH.A(161371);

	[CompilerGenerated]
	private static bool m_A;

	private static bool ShortcutManagerOpen
	{
		[CompilerGenerated]
		get
		{
			return ShortcutManager.m_A;
		}
		[CompilerGenerated]
		set
		{
			ShortcutManager.m_A = value;
		}
	}

	public static void Toggle(bool blnPressed)
	{
		try
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
						A();
						clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)13, VH.A(161338));
						return;
					}
				}
			}
			B();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A()
	{
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_004f: Expected O, but got Unknown
		ShortcutManagerOpen = true;
		ctpShortcuts2 control = new ctpShortcuts2();
		try
		{
			CustomTaskPane customTaskPane = MH.A.CustomTaskPanes.Add(control, ShortcutManager.m_A, MH.A.Application.ActiveWindow);
			try
			{
				customTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
				customTaskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
				clsDisplay val = new clsDisplay();
				customTaskPane.Width = checked((int)Math.Round(370.0 * val.X));
				val = null;
				customTaskPane.Visible = true;
				customTaskPane.VisibleChanged += A;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			customTaskPane = null;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
	}

	private static void B()
	{
		clsPanes.A(ShortcutManager.m_A);
		ShortcutManagerOpen = false;
		KH.A.InvalidateControl(ShortcutManager.m_B);
	}

	private static void A(object A, EventArgs B)
	{
		if ((A as CustomTaskPane).Visible)
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
			if (KH.A)
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
				ShortcutManager.B();
				return;
			}
		}
	}

	public static bool IsOpen()
	{
		return ShortcutManagerOpen;
	}

	public static void Refresh()
	{
		try
		{
			if (!IsOpen())
			{
				return;
			}
			while (true)
			{
				switch (7)
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
					((wpfShortcuts)((ctpShortcuts2)A().Control).ElementHost1.Child).PopulateList();
					return;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
					return;
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
	}

	private static CustomTaskPane A()
	{
		CustomTaskPaneCollection customTaskPanes = MH.A.CustomTaskPanes;
		checked
		{
			for (int i = customTaskPanes.Count - 1; i >= 0; i += -1)
			{
				if (Operators.CompareString(customTaskPanes[i].Title, ShortcutManager.m_A, TextCompare: false) != 0)
				{
					continue;
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
					return customTaskPanes[i];
				}
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				customTaskPanes = null;
				return null;
			}
		}
	}
}
