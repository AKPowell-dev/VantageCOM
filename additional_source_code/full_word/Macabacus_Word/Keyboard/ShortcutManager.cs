using System;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Keyboard;

public sealed class ShortcutManager
{
	private static readonly string m_A = XC.A(3286);

	private static readonly string m_B = XC.A(3319);

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
					switch (2)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						A();
						clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)13, XC.A(3286));
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
		//IL_0045: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		ShortcutManagerOpen = true;
		ctpShortcuts2 control = new ctpShortcuts2();
		try
		{
			CustomTaskPane customTaskPane = PC.A.CustomTaskPanes.Add(control, ShortcutManager.m_A, PC.A.Application.ActiveWindow);
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
		MC.A(ShortcutManager.m_A);
		ShortcutManagerOpen = false;
		NC.A.InvalidateControl(ShortcutManager.m_B);
	}

	private static void A(object A, EventArgs B)
	{
		if ((A as CustomTaskPane).Visible || NC.B)
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
			ShortcutManager.B();
			return;
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
					((wpfShortcuts)((ctpShortcuts2)MC.A(ShortcutManager.m_A).Control).ElementHost1.Child).PopulateList();
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
}
