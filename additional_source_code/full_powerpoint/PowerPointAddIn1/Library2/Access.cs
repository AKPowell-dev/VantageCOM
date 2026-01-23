using System;
using A;
using MacabacusMacros.Libraries;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Library2;

public sealed class Access
{
	public static void AfterPresentationOpen(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
	{
		int num;
		try
		{
			num = Pres.Windows.Count;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			num = 1;
			ProjectData.ClearProjectError();
		}
		if (Pres.ReadOnly != MsoTriState.msoFalse || num <= 0)
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
			if (Access.UserHasAccess(Pres.FullName))
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
				try
				{
					Pres.Close();
					Forms.WarningMessage(AH.A(67826));
					return;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
					return;
				}
			}
		}
	}
}
