using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.TextOps.Redaction.Redactors;

public sealed class ChartRedactor
{
	public static void RedactInlineChart(InlineShape shp)
	{
		PictureRedactor.RedactInlinePicture(shp);
	}

	public static Shape RedactFloatingChart(Shape shp)
	{
		PictureRedactor.RedactFloatingPicture(shp);
		return shp;
	}
}
