using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class HyperlinkItem : ExploreItem
{
	private bool A;

	[CompilerGenerated]
	private Hyperlink A;

	public override bool IsSelected
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(21693));
			Refresh();
		}
	}

	internal Hyperlink Hyperlink
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	public HyperlinkItem(WorksheetItem wsi, Hyperlink hyp)
		: base(wsi, Constants.ColorPalette.Purple.Clone(), Props.Icons.GeoHyperlink, 39)
	{
		Hyperlink = hyp;
		base.Range = hyp.Range;
		Refresh();
	}

	public override void Refresh()
	{
		((BaseItem)this).Label = Hyperlink.TextToDisplay;
		base.Tooltip = Hyperlink.Address;
	}

	public override void Delete()
	{
		if (MessageBox.Show(VH.A(116350), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
		{
			return;
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
			Hyperlink.Delete();
			base.Parent.A(this);
			return;
		}
	}

	public override void Search(string strQuery)
	{
		int isHighlighted;
		if (!((BaseItem)this).Label.ToLower().Contains(strQuery) && !Hyperlink.Address.ToLower().Contains(strQuery))
		{
			while (true)
			{
				switch (7)
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
			isHighlighted = ((Operators.CompareString(strQuery, VH.A(116443), TextCompare: false) == 0) ? 1 : 0);
		}
		else
		{
			isHighlighted = 1;
		}
		((BaseItem)this).IsHighlighted = (byte)isHighlighted != 0;
	}

	public void Edit()
	{
		Hyperlink.Application.CommandBars.ExecuteMso(VH.A(116464));
		System.Windows.Forms.Application.DoEvents();
	}

	public void Follow()
	{
		Hyperlink.Follow(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
	}
}
