using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class TableCellMargins : BaseError
{
	private List<CellPadding> A;

	private List<CellPadding> FixOptions
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
		}
	}

	public TableCellMargins(Table tbl, List<string> listLabels, List<Range> listRanges, List<CellPadding> listFixes)
		: base(ErrorType.TableCellMargins, ((Settings)Main.Analysis.Options).TableCellMargins, tbl, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		base.Ranges = listRanges;
		((BaseError)this).DisplayText = listLabels;
		FixOptions = listFixes;
		((BaseError)this).Title = XC.A(35373);
		((BaseError)this).Subtitle = XC.A(35424);
	}

	public override void FixAction(int i)
	{
		CellPadding cellPadding = FixOptions[i];
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(35340));
		foreach (Range range in base.Ranges)
		{
			Cell obj = (Cell)range.Parent;
			obj.TopPadding = cellPadding.Top;
			obj.RightPadding = cellPadding.Right;
			obj.BottomPadding = cellPadding.Bottom;
			obj.LeftPadding = cellPadding.Left;
			_ = null;
		}
		cellPadding = default(CellPadding);
		undoRecord.EndCustomRecord();
		undoRecord = null;
	}
}
