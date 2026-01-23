using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Fix;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingDataTableFontColor : BaseColorError
{
	public NonconformingDataTableFontColor(object shp, int intColor, DataTable tbl, Severity sev)
		: base(ErrorType.ColorPaletteChartDataTableFont, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.DataTable = tbl;
		((BaseError)this).Title = XC.A(31904);
		((BaseError)this).Subtitle = XC.A(31977);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		Macabacus_Word.Proofing.Fix.Colors.RecolorChartFont(base.DataTable.Format, ColorTranslator.ToOle(color));
		undoRecord.EndCustomRecord();
	}
}
