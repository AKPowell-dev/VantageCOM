using System.Reflection;
using System.Runtime.CompilerServices;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class FormatItem : SearchItem
{
	public FormatItem(WorksheetItem wsi, Range rng)
		: base(wsi, rng)
	{
	}

	public override void Refresh()
	{
		((BaseItem)this).Label = base.Range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
	}
}
