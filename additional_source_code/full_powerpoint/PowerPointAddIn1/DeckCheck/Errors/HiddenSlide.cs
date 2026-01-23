using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class HiddenSlide : BaseError
{
	public HiddenSlide(Slide sld)
		: base(ErrorType.HiddenSlide, Main.Analysis.Options.HiddenSlides, sld, null, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(36247);
		((BaseError)this).Subtitle = AH.A(36272) + sld.SlideIndex + AH.A(36285);
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		base.Slide.SlideShowTransition.Hidden = MsoTriState.msoFalse;
	}
}
