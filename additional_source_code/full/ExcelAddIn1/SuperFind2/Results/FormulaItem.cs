using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class FormulaItem : SearchItem
{
	public FormulaItem(WorksheetItem wsi, Range rng)
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
			text = text + VH.A(116343) + Strings.Mid(Conversions.ToString(NewLateBinding.LateGet(base.Range, null, VH.A(8714), new object[0], null, null, null)), 2);
		}
		((BaseItem)this).Label = text;
	}
}
