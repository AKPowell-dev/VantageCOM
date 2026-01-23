using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class HiddenSlides
{
	public void Check(Slide sld)
	{
		if (sld.SlideShowTransition.Hidden == MsoTriState.msoTrue)
		{
			Main.Analysis.Errors.Add(new HiddenSlide(sld));
		}
	}
}
