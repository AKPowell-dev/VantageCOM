using System.Collections.Generic;
using System.Runtime.CompilerServices;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Fix;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class HyphenSpacingImproper : BaseTextError
{
	public HyphenSpacingImproper(Range rng, List<string> listLabels, List<string> listFixes)
		: base(ErrorType.Text, (Severity)3, rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		BaseError val = (BaseError)(object)this;
		Errors.HyphenSpacingImproper(ref val, GenerateSnippet(rng));
	}

	public HyphenSpacingImproper(object shp, TextRange2 rng, List<string> listLabels, List<string> listFixes)
		: base(ErrorType.Text, (Severity)3, RuntimeHelpers.GetObjectValue(shp), rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		BaseError val = (BaseError)(object)this;
		Errors.HyphenSpacingImproper(ref val, GenerateSnippet(rng));
	}

	public override void FixAction(int i)
	{
		Text.ReplaceText(this, i);
	}
}
