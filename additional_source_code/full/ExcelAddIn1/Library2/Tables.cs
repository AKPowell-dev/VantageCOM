using ExcelAddIn1.Library2.UI;

namespace ExcelAddIn1.Library2;

public sealed class Tables
{
	public static void ShowInLibrary()
	{
		Pane.Show(blnShapes: false, blnImages: false, blnCharts: false, blnText: false, blnTables: true);
	}
}
