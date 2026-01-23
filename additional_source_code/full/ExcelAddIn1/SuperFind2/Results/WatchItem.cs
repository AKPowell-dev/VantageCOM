using System;
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

public sealed class WatchItem : ExploreItem
{
	private bool A;

	[CompilerGenerated]
	private Watch A;

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

	internal Watch Watch
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

	public WatchItem(WorksheetItem wsi, Watch watch)
		: base(wsi, Constants.ColorPalette.Brown.Clone(), Props.Icons.GeoWatch, 7)
	{
		Watch = watch;
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
		string text = (applicationLanguage - 1) switch
		{
			0 => VH.A(123696), 
			2 => VH.A(123707), 
			1 => VH.A(123728), 
			_ => VH.A(123696), 
		};
		try
		{
			base.Range = (Range)Watch.Source;
			((BaseItem)this).Label = text + VH.A(17350) + base.Range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public override void Delete()
	{
		if (MessageBox.Show(VH.A(123741), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
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
			try
			{
				Watch.Delete();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			base.Parent.A(this);
			return;
		}
	}

	public override void Search(string strQuery)
	{
		int isHighlighted;
		if (!((BaseItem)this).Label.ToLower().Contains(strQuery))
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
			isHighlighted = ((Operators.CompareString(strQuery, VH.A(123828), TextCompare: false) == 0) ? 1 : 0);
		}
		else
		{
			isHighlighted = 1;
		}
		((BaseItem)this).IsHighlighted = (byte)isHighlighted != 0;
	}
}
