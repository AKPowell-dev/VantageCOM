using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class MissingOpeningSquareBracket : BaseTextError
{
	public MissingOpeningSquareBracket(Range rng)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).PunctuationMissing, rng, blnHasFix: false)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		A();
		((BaseError)this).Subtitle = GenerateSnippet(rng);
	}

	public MissingOpeningSquareBracket(object shp, TextRange2 rng)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).PunctuationMissing, RuntimeHelpers.GetObjectValue(shp), rng, blnHasFix: false)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		A();
		((BaseError)this).Subtitle = GenerateSnippet(rng);
	}

	private void A()
	{
		((BaseError)this).Title = XC.A(35882);
	}
}
