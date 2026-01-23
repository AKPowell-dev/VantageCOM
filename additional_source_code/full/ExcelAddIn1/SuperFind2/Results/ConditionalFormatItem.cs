using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class ConditionalFormatItem : ExploreItem
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

	public ConditionalFormatItem(WorksheetItem wsi, Range rng)
		: base(wsi, Constants.ColorPalette.DeepPurple.Clone(), Props.Icons.GeoConditional, 5)
	{
		base.Range = rng;
		Refresh();
	}

	public override void Refresh()
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_0005: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000a: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Expected I4, but got Unknown
		Language applicationLanguage = clsEnvironment.ApplicationLanguage;
		((BaseItem)this).Label = (applicationLanguage - 1) switch
		{
			0 => VH.A(116007), 
			2 => VH.A(116044), 
			1 => VH.A(116083), 
			_ => VH.A(116007), 
		} + VH.A(17350) + base.Range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
	}

	public override void Delete()
	{
		if (MessageBox.Show(VH.A(116140), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
		{
			base.Range.FormatConditions.Delete();
			base.Parent.A(this);
		}
	}

	public override void Search(string strQuery)
	{
		int isHighlighted;
		if (!((BaseItem)this).Label.ToLower().Contains(strQuery))
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
			isHighlighted = ((Operators.CompareString(strQuery, VH.A(116261), TextCompare: false) == 0) ? 1 : 0);
		}
		else
		{
			isHighlighted = 1;
		}
		((BaseItem)this).IsHighlighted = (byte)isHighlighted != 0;
	}

	public void Modify()
	{
		base.Range.Application.CommandBars.ExecuteMso(VH.A(116286));
		System.Windows.Forms.Application.DoEvents();
	}
}
