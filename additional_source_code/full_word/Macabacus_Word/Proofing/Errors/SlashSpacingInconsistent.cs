using System.Runtime.CompilerServices;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Fix;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class SlashSpacingInconsistent : BaseTextError
{
	public SlashSpacingInconsistent(Range rng, string strFix)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).PunctuationSpacingInconsistent, rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		A(strFix);
	}

	public SlashSpacingInconsistent(object shp, TextRange2 rng, string strFix)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).PunctuationSpacingInconsistent, RuntimeHelpers.GetObjectValue(shp), rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		A(strFix);
	}

	private void A(string A)
	{
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.SlashSpacingInconsistent(ref val, ((Settings)Main.Analysis.Options).SlashSpacingConvention);
	}

	public override void FixAction(int i)
	{
		Text.ReplaceText(this, i);
	}
}
