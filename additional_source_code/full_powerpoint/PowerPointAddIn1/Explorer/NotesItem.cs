using System.Windows.Forms;
using System.Windows.Media;
using A;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Explorer;

public sealed class NotesItem : ContentItem
{
	private new bool m_A;

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
			A();
		}
	}

	public NotesItem(SlideItem si, Shape shp, SolidColorBrush brush)
		: base(si, "", brush, Pane.CachedObjects.GeoNotes)
	{
		base.Shape = shp;
		A();
		SearchOnInstantiate();
	}

	public override void Refresh()
	{
		A();
		base.PreviewImage = null;
	}

	public override void Delete()
	{
		if (MessageBox.Show(AH.A(114450), AH.A(5874), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
		{
			return;
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			base.Shape.TextFrame2.DeleteText();
			base.Parent.RemoveChild(this);
			return;
		}
	}

	public override void Search(string strQuery)
	{
		int isHighlighted;
		if (!base.Tooltip.ToLower().Contains(strQuery))
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
			isHighlighted = ((Operators.CompareString(strQuery, AH.A(114555), TextCompare: false) == 0) ? 1 : 0);
		}
		else
		{
			isHighlighted = 1;
		}
		((BaseItem)this).IsHighlighted = (byte)isHighlighted != 0;
	}

	private void A()
	{
		string text = (base.Tooltip = base.Shape.TextFrame2.TextRange.Text.Trim());
		text = text.Replace(AH.A(7894), AH.A(14625)).Replace(AH.A(47331), AH.A(14625)).Replace(AH.A(47334), AH.A(14625));
		if (text.Length > 97)
		{
			while (true)
			{
				switch (5)
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
			text = text.Substring(0, 96) + AH.A(114568);
		}
		((BaseItem)this).Label = text;
	}
}
