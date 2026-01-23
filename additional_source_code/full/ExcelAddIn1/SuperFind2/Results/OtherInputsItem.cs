using System;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros.Explorer;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class OtherInputsItem : ExploreItem
{
	public OtherInputsItem(WorksheetItem wsi, string strLabel)
		: base(wsi, Constants.ColorPalette.Inputs.Clone(), Props.Icons.GeoNumbers, 41)
	{
		((BaseItem)this).Label = strLabel;
	}

	public override void Refresh()
	{
		throw new NotImplementedException();
	}

	public override void Delete()
	{
		throw new NotImplementedException();
	}

	public override void Search(string strQuery)
	{
	}
}
