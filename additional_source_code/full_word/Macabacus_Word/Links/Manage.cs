using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Links;

public sealed class Manage
{
	public static void LinkWizard()
	{
		if (!Access.AllowWordOperation((PlanType)5, (Restriction)2, false))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			bool flag = false;
			try
			{
				IEnumerable<wpfManageLinks> source = Application.Current.Windows.OfType<wpfManageLinks>();
				if (source.Any())
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
					source.ElementAt(0).Activate();
					flag = true;
				}
				source = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (!flag)
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
				wpfManageLinks wpfManageLinks2 = new wpfManageLinks();
				if (Properties.ManageLinksHeight > 0.0)
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
					wpfManageLinks2.Height = Properties.ManageLinksHeight;
				}
				if (Properties.ManageLinksWidth > 0.0)
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
					wpfManageLinks2.Width = Properties.ManageLinksWidth;
				}
				wpfManageLinks2.Show();
				wpfManageLinks2 = null;
			}
			clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)10, XC.A(14450));
			return;
		}
	}
}
