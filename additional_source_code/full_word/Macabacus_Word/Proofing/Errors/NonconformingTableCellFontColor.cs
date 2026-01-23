using System.Collections.Generic;
using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingTableCellFontColor : BaseColorError
{
	public NonconformingTableCellFontColor(Table tbl, int intColor, List<Range> listRanges, Severity sev)
		: base(ErrorType.ColorPaletteFont, sev, tbl, intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		base.Ranges = listRanges;
		((BaseError)this).Title = XC.A(32519);
		((BaseError)this).Subtitle = XC.A(32570);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		int rGB = ColorTranslator.ToOle(color);
		using (List<Range>.Enumerator enumerator = base.Ranges.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				enumerator.Current.Font.TextColor.RGB = rGB;
			}
			while (true)
			{
				switch (1)
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
