using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class SpillRangeItem : ExploreItem
{
	private bool A;

	public override bool IsSelected
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(21693));
			Refresh();
		}
	}

	public SpillRangeItem(WorksheetItem wsi, Range rng)
		: base(wsi, Constants.ColorPalette.GoldenRod.Clone(), Props.Icons.GeoInkSpill, 30)
	{
		base.Range = rng;
		Refresh();
	}

	public override void Refresh()
	{
		((BaseItem)this).Label = VH.A(118088) + base.Range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
	}

	public override void Delete()
	{
		throw new NotImplementedException();
	}

	public override void Search(string strQuery)
	{
		((BaseItem)this).IsHighlighted = ((BaseItem)this).Label.ToLower().Contains(strQuery) || Operators.CompareString(strQuery, VH.A(118117), TextCompare: false) == 0;
	}
}
