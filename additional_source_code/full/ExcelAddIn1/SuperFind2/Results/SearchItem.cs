using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Controls;
using A;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public abstract class SearchItem : ResultItem
{
	private ControlTemplate A;

	public ControlTemplate RangeIcon
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(120220));
		}
	}

	public SearchItem(WorksheetItem wsi, Range rng)
		: base(wsi, rng.Worksheet, rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), Props.Icons.GeoUsedRange, 0)
	{
		A = null;
		base.Range = rng;
		ControlTemplate rangeIcon;
		if (!Operators.ConditionalCompareObjectEqual(rng.Cells.CountLarge, 1, TextCompare: false))
		{
			while (true)
			{
				switch (1)
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
			rangeIcon = Props.MultiCellIcon;
		}
		else
		{
			rangeIcon = Props.SingleCellIcon;
		}
		RangeIcon = rangeIcon;
		base.FontColor = wsi.FontColor;
	}

	public abstract void Refresh();
}
