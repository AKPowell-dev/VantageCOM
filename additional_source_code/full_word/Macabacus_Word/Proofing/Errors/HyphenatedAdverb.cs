using System.Runtime.CompilerServices;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Fix;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class HyphenatedAdverb : BaseTextError
{
	public HyphenatedAdverb(Range rng, string strFix)
		: base(ErrorType.Text, (Severity)3, rng, blnHasFix: true)
	{
		BaseError val = (BaseError)(object)this;
		Errors.HyphenAdverb(ref val, GenerateSnippet(rng));
	}

	public HyphenatedAdverb(object shp, TextRange2 rng, string strFix)
		: base(ErrorType.Text, (Severity)3, RuntimeHelpers.GetObjectValue(shp), rng, blnHasFix: true)
	{
		BaseError val = (BaseError)(object)this;
		Errors.HyphenAdverb(ref val, GenerateSnippet(rng));
	}

	public override void FixAction(int i)
	{
		Text.ReplaceText(this, i);
	}
}
