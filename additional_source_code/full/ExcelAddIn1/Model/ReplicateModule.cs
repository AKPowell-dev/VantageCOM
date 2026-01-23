using System;
using System.Collections;
using A;
using ExcelAddIn1.Audit.Visualizations;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Model;

public sealed class ReplicateModule
{
	public static void Go()
	{
		if (!Access.AllowExcelOperation((PlanType)5, (Restriction)2, false))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		bool flag = default(bool);
		bool flag2 = default(bool);
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
			Application application = MH.A.Application;
			Common.ClearVisualizations(application);
			if (application.Selection is Range)
			{
				if (((Range)application.Selection).Areas.Count == 1)
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
						enumerator = application.ActiveWorkbook.Worksheets.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Worksheet worksheet = (Worksheet)enumerator.Current;
							if (Strings.InStr(worksheet.Name, VH.A(90521)) > 0)
							{
								flag = true;
								continue;
							}
							if (Strings.InStr(worksheet.Name, VH.A(90536)) <= 0)
							{
								continue;
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								break;
							}
							flag2 = true;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								(enumerator as IDisposable).Dispose();
								break;
							}
						}
					}
					string text;
					if (!flag)
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
						text = VH.A(90547);
					}
					else if (!flag2)
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
						text = VH.A(224);
					}
					else
					{
						text = VH.A(90560);
					}
					if (application.ActiveWindow.SelectedSheets.Count > 1)
					{
						NewLateBinding.LateCall(application.ActiveSheet, null, VH.A(51162), new object[0], null, null, null, IgnoreReturn: true);
					}
					wpfReplicate wpfReplicate2 = new wpfReplicate();
					wpfReplicate2.cbxBaseName.Text = text;
					wpfReplicate2.ShowDialog();
					if (wpfReplicate2.DialogResult.HasValue && wpfReplicate2.DialogResult.Value)
					{
						clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(90577));
					}
					wpfReplicate2 = null;
				}
				else
				{
					Forms.WarningMessage(VH.A(90610));
				}
			}
			else
			{
				Forms.WarningMessage(VH.A(90717));
			}
			application = null;
			return;
		}
	}
}
