using System;
using A;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros.Explorer;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class OtherNamesItem : ExploreItem
{
	public OtherNamesItem(WorksheetItem wsi, string strLabel)
		: base(wsi, Constants.ColorPalette.LightGreen.Clone(), Props.Icons.GeoName, 36)
	{
		((BaseItem)this).Label = strLabel;
		base.Tooltip = VH.A(115506);
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
		int isHighlighted;
		if (!((BaseItem)this).Label.ToLower().Contains(strQuery))
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
			isHighlighted = ((Operators.CompareString(strQuery, VH.A(115495), TextCompare: false) == 0) ? 1 : 0);
		}
		else
		{
			isHighlighted = 1;
		}
		((BaseItem)this).IsHighlighted = (byte)isHighlighted != 0;
	}
}
