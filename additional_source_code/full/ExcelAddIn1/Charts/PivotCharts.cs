using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class PivotCharts
{
	public static void HideFieldButtonsAll()
	{
		A(A: false);
	}

	public static void ShowFieldButtonsAll()
	{
		A(A: true);
	}

	public static void HideFieldButtonsSheet()
	{
		B(A: false);
	}

	public static void ShowFieldButtonsSheet()
	{
		B(A: true);
	}

	private static void A(bool A)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = MH.A.Application.ActiveWorkbook.Worksheets.GetEnumerator();
			while (enumerator.MoveNext())
			{
				PivotCharts.A((Worksheet)enumerator.Current, A);
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
				return;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	private static void B(bool A)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		if (application.ActiveSheet is Worksheet)
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
			PivotCharts.A((Worksheet)application.ActiveSheet, A);
		}
		application = null;
	}

	private static void A(Worksheet A, bool B)
	{
		try
		{
			ChartObjects chartObjects = (ChartObjects)A.ChartObjects(RuntimeHelpers.GetObjectValue(Missing.Value));
			int count = chartObjects.Count;
			for (int i = 1; i <= count; i = checked(i + 1))
			{
				Chart chart = ((ChartObject)chartObjects.Item(i)).Chart;
				if (chart.PivotLayout != null)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					chart.ShowAllFieldButtons = B;
				}
				chart = null;
			}
			chartObjects = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			MessageBox.Show(ex2.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
			ProjectData.ClearProjectError();
		}
	}
}
