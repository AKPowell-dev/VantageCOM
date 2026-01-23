using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Fix;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class SpellingCanceled : BaseTextError
{
	public SpellingCanceled(Range rng, string strFix, Severity sev)
		: base(ErrorType.Text, sev, rng, blnHasFix: true)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.SpellingCanceled(ref val, strFix);
	}

	public override void FixAction(int i)
	{
		Text.ReplaceText(this, i);
	}
}
