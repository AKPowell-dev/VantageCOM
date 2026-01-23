using System.Runtime.CompilerServices;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Fix;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class AbbreviationSpacing : BaseTextError
{
	public AbbreviationSpacing(Range rng, string strFix, int intRequiredSpaces)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).MillionsBillionsAbbreviation, rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		A(strFix, intRequiredSpaces);
	}

	public AbbreviationSpacing(object shp, TextRange2 rng, string strFix, int intRequiredSpaces)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).MillionsBillionsAbbreviation, RuntimeHelpers.GetObjectValue(shp), rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		A(strFix, intRequiredSpaces);
	}

	private void A(string A, int B)
	{
		BaseError val = (BaseError)(object)this;
		Errors.AbbreviationSpacing(ref val, B);
	}

	public override void FixAction(int i)
	{
		Text.ReplaceText(this, i);
	}
}
