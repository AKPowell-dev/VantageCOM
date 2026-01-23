using System.Collections.Generic;
using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingTableCellFillColor : BaseColorError
{
	public NonconformingTableCellFillColor(Table tbl, int intColor, List<Range> listRanges, Severity sev)
		: base(ErrorType.ColorPaletteFill, sev, tbl, intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		base.Ranges = listRanges;
		((BaseError)this).Title = XC.A(32326);
		((BaseError)this).Subtitle = XC.A(32377);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		ColorTranslator.ToOle(color);
		using (List<Range>.Enumerator enumerator = base.Ranges.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				_ = enumerator.Current;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				break;
			}
		}
		undoRecord.EndCustomRecord();
		undoRecord = null;
	}
}
