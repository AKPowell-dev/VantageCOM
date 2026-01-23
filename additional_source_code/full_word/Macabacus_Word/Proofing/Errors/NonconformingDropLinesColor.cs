using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingDropLinesColor : BaseColorError
{
	public NonconformingDropLinesColor(object shp, int intColor, ChartGroup grp, Severity sev)
		: base(ErrorType.ColorPaletteChartDropLines, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.ChartGroup = grp;
		((BaseError)this).Title = XC.A(32125);
		((BaseError)this).Subtitle = XC.A(32188);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		base.ChartGroup.DropLines.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
		undoRecord.EndCustomRecord();
	}
}
