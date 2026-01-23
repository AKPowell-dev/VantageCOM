using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.Presentation;

public sealed class Helpers
{
	public static Microsoft.Office.Interop.PowerPoint.Presentation OpenQuietly(Application ppApp, string strPath)
	{
		return ppApp.Presentations.Open(strPath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
	}

	public static void CloseQuietly(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = pres;
		if (presentation.ReadOnly == MsoTriState.msoFalse)
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
			presentation.Saved = MsoTriState.msoTrue;
		}
		presentation.Close();
		presentation = null;
	}
}
