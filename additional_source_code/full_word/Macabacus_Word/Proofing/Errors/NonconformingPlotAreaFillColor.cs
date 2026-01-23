using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingPlotAreaFillColor : BaseColorError
{
	public NonconformingPlotAreaFillColor(object shp, int intColor, PlotArea plot, Severity sev)
		: base(ErrorType.ColorPaletteChartPlotFill, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.PlotArea = plot;
		((BaseError)this).Title = XC.A(34255);
		((BaseError)this).Subtitle = XC.A(34326);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		base.PlotArea.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
		undoRecord.EndCustomRecord();
	}
}
