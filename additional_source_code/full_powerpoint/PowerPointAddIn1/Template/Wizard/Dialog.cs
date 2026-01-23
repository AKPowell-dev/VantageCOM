using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Template.Wizard;

public sealed class Dialog
{
	public static void Show()
	{
		if (!Licensing.AllowTemplateOperation())
		{
			return;
		}
		bool flag = false;
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation;
		try
		{
			activePresentation = NG.A.Application.ActivePresentation;
			if (Operators.CompareString(Path.GetExtension(activePresentation.Name), AH.A(69996), TextCompare: false) != 0)
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
				if (Forms.OkCancelMessage2(AH.A(121845)) == DialogResult.Cancel)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
			}
			try
			{
				IEnumerable<wpfWizard> source = System.Windows.Application.Current.Windows.OfType<wpfWizard>();
				if (source.Any())
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
				new wpfWizard(activePresentation.FullName).Show();
				_ = null;
			}
			clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)11, AH.A(122397));
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		activePresentation = null;
	}
}
