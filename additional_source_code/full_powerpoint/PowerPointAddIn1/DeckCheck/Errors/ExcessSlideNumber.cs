using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ExcessSlideNumber : BaseError
{
	public ExcessSlideNumber(Slide sld, string strSubtitle)
		: base(ErrorType.ExcessSlideNumber, Main.Analysis.Options.SlideNumbers, sld, null, blnHasFix: false)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(36066);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = AH.A(36107);
	}
}
