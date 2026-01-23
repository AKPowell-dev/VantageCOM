using System.Runtime.CompilerServices;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class DummyText : BaseTextError
{
	public DummyText(Range rng)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).DummyText, rng, blnHasFix: false)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.DummyText(ref val, GenerateSnippet(rng));
	}

	public DummyText(object shp, TextRange2 rng)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).DummyText, RuntimeHelpers.GetObjectValue(shp), rng, blnHasFix: false)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.DummyText(ref val, GenerateSnippet(rng));
	}
}
