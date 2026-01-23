using A;
using MacabacusMacros.Proofing;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class MultipleSlideMasters : BaseError
{
	public MultipleSlideMasters(string strSubtitle)
		: base(ErrorType.MultipleMasters, Main.Analysis.Options.MultipleSlideMasters, null, null, blnHasFix: false)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(37713);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = AH.A(37758);
	}
}
