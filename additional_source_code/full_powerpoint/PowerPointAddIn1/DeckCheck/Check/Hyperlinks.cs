using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class Hyperlinks
{
	public void Check(Slide sld)
	{
		int count = sld.Hyperlinks.Count;
		if (count <= 0)
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.Hyperlinks(sld, count));
			return;
		}
	}
}
