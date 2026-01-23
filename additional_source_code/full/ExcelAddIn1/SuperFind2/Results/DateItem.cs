using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class DateItem : SearchItem
{
	public DateItem(WorksheetItem wsi, Range rng)
		: base(wsi, rng)
	{
		Refresh();
	}

	public override void Refresh()
	{
		string text = base.Range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		if (Operators.ConditionalCompareObjectEqual(base.Range.Cells.CountLarge, 1, TextCompare: false))
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
			string text2 = base.Range.Text.ToString().Trim();
			if (text2.Length > 0)
			{
				text = text + VH.A(116343) + text2;
			}
		}
		((BaseItem)this).Label = text;
	}
}
