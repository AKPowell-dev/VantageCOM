using System;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Media;
using A;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Links;

namespace PowerPointAddIn1.Explorer;

public sealed class HyperlinkItem : ContentItem
{
	private new bool m_A;

	[CompilerGenerated]
	private new Hyperlink m_A;

	public override bool IsSelected
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(62846));
			Refresh();
		}
	}

	public Hyperlink Hyperlink
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public HyperlinkItem(SlideItem si, Hyperlink hyp, SolidColorBrush brush)
		: base(si, "", brush, Pane.CachedObjects.GeoHyperlink)
	{
		Hyperlink = hyp;
		A();
		base.IsLinked = PowerPointAddIn1.Links.Hyperlinks.IsLinked(hyp);
		base.Tooltip = hyp.Address;
		SearchOnInstantiate();
	}

	public override void Refresh()
	{
		A();
	}

	public override void Delete()
	{
		if (MessageBox.Show(AH.A(113760), AH.A(5874), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
		{
			PowerPointAddIn1.Links.Hyperlinks.C(Hyperlink);
			base.Parent.RemoveChild(this);
		}
	}

	public override void Search(string strQuery)
	{
		int isHighlighted;
		if (!((BaseItem)this).Label.ToLower().Contains(strQuery))
		{
			while (true)
			{
				switch (6)
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
			if (!Hyperlink.Address.ToLower().Contains(strQuery))
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					break;
				}
				isHighlighted = ((Operators.CompareString(strQuery, AH.A(113853), TextCompare: false) == 0) ? 1 : 0);
				goto IL_0070;
			}
		}
		isHighlighted = 1;
		goto IL_0070;
		IL_0070:
		((BaseItem)this).IsHighlighted = (byte)isHighlighted != 0;
	}

	public void Edit()
	{
		Hyperlink.Application.CommandBars.ExecuteMso(AH.A(113874));
		Hyperlink.Application.CommandBars.ExecuteMso(AH.A(113901));
		Hyperlink.Application.CommandBars.ExecuteMso(AH.A(113932));
		System.Windows.Forms.Application.DoEvents();
	}

	public void Follow()
	{
		Hyperlink.Follow();
	}

	private void A()
	{
		try
		{
			((BaseItem)this).Label = Hyperlink.TextToDisplay.Trim();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			((BaseItem)this).Label = PowerPointAddIn1.Links.Hyperlinks.GetParentShape(Hyperlink, blnIgnoreTables: true).Name;
			ProjectData.ClearProjectError();
		}
	}
}
