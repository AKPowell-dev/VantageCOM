using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingDataTableFillColor : BaseColorError
{
	public NonconformingDataTableFillColor(object shp, int intColor, DataTable tbl, Severity sev)
		: base(ErrorType.ColorPaletteChartDataTableFill, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.DataTable = tbl;
		((BaseError)this).Title = XC.A(31683);
		((BaseError)this).Subtitle = XC.A(31756);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		base.DataTable.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
		undoRecord.EndCustomRecord();
	}
}
