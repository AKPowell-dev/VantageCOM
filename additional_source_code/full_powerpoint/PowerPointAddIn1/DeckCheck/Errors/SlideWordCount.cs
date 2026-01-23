using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class SlideWordCount : BaseError
{
	public SlideWordCount(Slide sld, int intCount, int intMax)
		: base(ErrorType.SlideWordCount, Main.Analysis.Options.SlideWordCount, sld, null, blnHasFix: false)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(37400);
		((BaseError)this).Subtitle = AH.A(37433) + intCount + AH.A(37464) + intMax;
	}
}
