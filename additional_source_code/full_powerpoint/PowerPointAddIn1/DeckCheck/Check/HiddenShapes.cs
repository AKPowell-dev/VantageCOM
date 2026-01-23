using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class HiddenShapes
{
	public void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		if (shp.Visible != MsoTriState.msoFalse)
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
			Main.Analysis.Errors.Add(new HiddenShape(sld, shp));
			return;
		}
	}
}
