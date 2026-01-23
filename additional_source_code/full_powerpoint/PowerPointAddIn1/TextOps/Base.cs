using MacabacusMacros;

namespace PowerPointAddIn1.TextOps;

public sealed class Base
{
	public static void LogActivity(string strActivity)
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)1, strActivity);
	}
}
