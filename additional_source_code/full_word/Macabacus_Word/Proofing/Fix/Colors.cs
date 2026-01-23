using System.Drawing;
using A;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Fix;

public sealed class Colors
{
	public static void ReplaceColor(BaseError err, Color clr)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		int rGB = ColorTranslator.ToOle(clr);
		if (err.InlineShape != null)
		{
			switch (err.Type)
			{
			case ErrorType.ColorPaletteFill:
				err.InlineShape.Fill.ForeColor.RGB = rGB;
				break;
			case ErrorType.ColorPaletteFont:
				err.InlineShape.Range.Font.Fill.ForeColor.RGB = rGB;
				break;
			case ErrorType.ColorPaletteBorder:
				err.InlineShape.Line.ForeColor.RGB = rGB;
				break;
			}
		}
		else if (err.Shape != null)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			switch (err.Type)
			{
			case ErrorType.ColorPaletteFill:
				err.Shape.Fill.ForeColor.RGB = rGB;
				break;
			case ErrorType.ColorPaletteFont:
				err.Shape.TextFrame2.TextRange.get_Characters(-1, -1).Font.Fill.ForeColor.RGB = rGB;
				break;
			case ErrorType.ColorPaletteBorder:
				err.Shape.Line.ForeColor.RGB = rGB;
				break;
			}
		}
		undoRecord.EndCustomRecord();
		undoRecord = null;
	}

	public static void RecolorChartFont(ChartFormat format, int intColor)
	{
		format.TextFrame2.TextRange.get_Characters(-1, -1).Font.Fill.ForeColor.RGB = intColor;
	}
}
