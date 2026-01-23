using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingPlotAreaBorderColor : BaseColorError
{
	public NonconformingPlotAreaBorderColor(object shp, int intColor, PlotArea plot, Severity sev)
		: base(ErrorType.ColorPaletteChartPlotBorder, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.PlotArea = plot;
		((BaseError)this).Title = XC.A(34030);
		((BaseError)this).Subtitle = XC.A(34105);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		base.PlotArea.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
		undoRecord.EndCustomRecord();
	}
}
