using System.Collections.Generic;
using System.Runtime.CompilerServices;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing;

public class BaseTextError : BaseError
{
	public BaseTextError(ErrorType errType, Severity sev, Range rng, bool blnHasFix, bool blnCanFixMultiple = false)
		: base(errType, sev, rng, blnHasFix, blnCanFixMultiple)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		base.Ranges = new List<Range>(new Range[1] { rng });
	}

	public BaseTextError(ErrorType errType, Severity sev, object shp, TextRange2 rng, bool blnHasFix, bool blnCanFixMultiple = false)
		: base(errType, sev, RuntimeHelpers.GetObjectValue(shp), blnHasFix, blnCanFixMultiple)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).TextRanges = new List<TextRange2>(new TextRange2[1] { rng });
	}
}
