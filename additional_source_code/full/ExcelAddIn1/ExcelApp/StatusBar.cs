using System.Timers;
using A;

namespace ExcelAddIn1.ExcelApp;

public sealed class StatusBar
{
	public static void SetText(string msg)
	{
		KH.A = new Timer(3000.0);
		KH.A.Elapsed += A;
		KH.A.AutoReset = false;
		KH.A.Enabled = true;
		MH.A.Application.StatusBar = msg;
	}

	private static void A(object A, ElapsedEventArgs B)
	{
		MH.A.Application.StatusBar = false;
		KH.A = null;
		KH.A.Dispose();
	}
}
