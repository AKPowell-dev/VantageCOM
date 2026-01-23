using System.Runtime.CompilerServices;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Fix;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class AbbreviationMillions : BaseTextError
{
	public AbbreviationMillions(Range rng, string strFix)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).MillionsBillionsAbbreviation, rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		A(strFix);
	}

	public AbbreviationMillions(object shp, TextRange2 rng, string strFix)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).MillionsBillionsAbbreviation, RuntimeHelpers.GetObjectValue(shp), rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		A(strFix);
	}

	private void A(string A)
	{
		BaseError val = (BaseError)(object)this;
		Errors.AbbreviationMillions(ref val, A);
	}

	public override void FixAction(int i)
	{
		Text.ReplaceText(this, i);
	}
}
