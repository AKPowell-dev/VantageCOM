using System;
using System.Collections.Generic;
using A;
using MacabacusMacros;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word;

public sealed class clsPanes
{
	private static CustomTaskPane A(Dictionary<int, CustomTaskPane> A)
	{
		CustomTaskPane value = null;
		try
		{
			A?.TryGetValue(PC.A.Application.ActiveWindow.Hwnd, out value);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return value;
	}

	public static bool IsVisible(Dictionary<int, CustomTaskPane> panes)
	{
		bool result;
		try
		{
			result = A(panes).Visible;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	internal static void A(CustomTaskPane A, clsDisplay B, int C = 0)
	{
		if (C == 0)
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
			C = N.Settings.TaskPaneWidth;
		}
		try
		{
			A.Width = checked((int)Math.Round((double)C * B.X));
		}
		catch (ArgumentException ex)
		{
			ProjectData.SetProjectError(ex);
			ArgumentException ex2 = ex;
			A.Width = 430;
			ProjectData.ClearProjectError();
		}
	}
}
