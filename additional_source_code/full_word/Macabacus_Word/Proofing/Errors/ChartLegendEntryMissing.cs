using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class ChartLegendEntryMissing : BaseError
{
	public ChartLegendEntryMissing(Shape shp, string strSubtitle, Legend legend)
		: base(ErrorType.ChartLegendEntryMissing, ((Settings)Main.Analysis.Options).CheckChartElements, shp, blnHasFix: true)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		base.Legend = legend;
		((BaseError)this).Title = XC.A(27534);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = XC.A(27597);
	}

	public override void FixAction()
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27501));
		Chart chart = base.Chart;
		XlLegendPosition position = chart.Legend.Position;
		chart.HasLegend = false;
		chart.HasLegend = true;
		chart.Legend.Position = position;
		_ = null;
		undoRecord.EndCustomRecord();
	}
}
