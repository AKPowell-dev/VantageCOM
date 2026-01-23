using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.TextOps.Redaction.Redactors;

public sealed class AutoShapeRedactor
{
	public static void RedactEntireRangeInAutoshape(Shape shp)
	{
		TextRedactor.RedactWordsInAutoshape(shp, shp.TextFrame.TextRange);
	}

	public static void RedactRangeAutoshape(Shape shp, Range rng)
	{
		TextRedactor.RedactWordsInAutoshape(shp, rng);
	}
}
