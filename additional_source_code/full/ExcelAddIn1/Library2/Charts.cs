using System;
using System.IO;
using A;
using ExcelAddIn1.Library2.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Library2;

public sealed class Charts
{
	public static void RetainChartTheme(Chart chtSource, Chart chtTarget)
	{
		try
		{
			string text = Path.Combine(Interaction.Environ(VH.A(86342)), VH.A(86351));
			chtSource.SaveChartTemplate(text);
			chtTarget.ApplyChartTemplate(text);
			File.Delete(text);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void ShowInLibrary()
	{
		ExcelAddIn1.Library2.UI.Pane.Show(blnShapes: false, blnImages: false, blnCharts: true, blnText: false, blnTables: false);
	}
}
