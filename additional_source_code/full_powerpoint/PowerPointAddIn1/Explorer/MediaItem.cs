using System.Windows.Forms;
using System.Windows.Media;
using A;
using MacabacusMacros.Explorer;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Explorer;

public sealed class MediaItem : ContentItem
{
	private new bool A;

	public override bool IsSelected
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(62846));
			Refresh();
		}
	}

	public MediaItem(SlideItem wsi, Microsoft.Office.Interop.PowerPoint.Shape shp, SolidColorBrush brush)
		: base(wsi, shp.Name, brush, Pane.CachedObjects.GeoMedia)
	{
		base.Shape = shp;
		base.IsLibraryContent = A(shp);
		UpdateColors(shp.Visible);
		Refresh();
		SearchOnInstantiate();
		if (shp.Type != MsoShapeType.msoMedia)
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (shp.MediaType != PpMediaType.ppMediaTypeSound)
			{
				return;
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				((BaseItem)this).Icon = Geometry.Parse(AH.A(114018));
				return;
			}
		}
	}

	public override void Refresh()
	{
		((BaseItem)this).Label = base.Shape.Name;
	}

	public override void Delete()
	{
		if (MessageBox.Show(AH.A(113231), AH.A(5874), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
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
			base.Shape.Delete();
			base.Parent.RemoveChild(this);
			return;
		}
	}

	public override void Search(string strQuery)
	{
		int isHighlighted;
		if (!((BaseItem)this).Label.ToLower().Contains(strQuery) && Operators.CompareString(strQuery, AH.A(114005), TextCompare: false) != 0)
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
			if (Operators.CompareString(strQuery, AH.A(113342), TextCompare: false) == 0)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					break;
				}
				isHighlighted = (base.IsLibraryContent ? 1 : 0);
			}
			else
			{
				isHighlighted = 0;
			}
		}
		else
		{
			isHighlighted = 1;
		}
		((BaseItem)this).IsHighlighted = (byte)isHighlighted != 0;
	}
}
