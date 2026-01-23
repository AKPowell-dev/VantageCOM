using A;
using MacabacusMacros.Proofing;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class SlideCount : BaseError
{
	public SlideCount(int intCount, int intMax)
		: base(ErrorType.SlideCount, Main.Analysis.Options.SlideCount, null, null, blnHasFix: false)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(36592);
		((BaseError)this).Subtitle = AH.A(36615) + intCount + AH.A(36660) + intMax + AH.A(14417);
	}
}
