using System.Windows.Forms;
using A;
using MacabacusMacros;

namespace ExcelAddIn1.RowsColumns;

public sealed class Core
{
	public static bool ConfirmMultipleSheets()
	{
		return MessageBox.Show(VH.A(170214), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK;
	}

	public static void LogActivity(string strActivity)
	{
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, strActivity);
	}
}
