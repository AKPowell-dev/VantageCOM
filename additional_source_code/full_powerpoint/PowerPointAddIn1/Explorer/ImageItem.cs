using System.Windows.Forms;
using System.Windows.Media;
using A;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Links;

namespace PowerPointAddIn1.Explorer;

public sealed class ImageItem : ContentItem
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

	public ImageItem(SlideItem wsi, Shape shp, SolidColorBrush brush)
		: base(wsi, shp.Name, brush, Pane.CachedObjects.GeoImage)
	{
		base.Shape = shp;
		base.IsLinked = PowerPointAddIn1.Links.Shapes.IsLinked(shp);
		base.IsLibraryContent = A(shp);
		UpdateColors(shp.Visible);
		Refresh();
		SearchOnInstantiate();
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
			switch (1)
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
		if (!((BaseItem)this).Label.ToLower().Contains(strQuery))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (Operators.CompareString(strQuery, AH.A(113983), TextCompare: false) != 0)
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
				if (Operators.CompareString(strQuery, AH.A(113331), TextCompare: false) == 0)
				{
					if (base.IsLinked)
					{
						goto IL_0096;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						break;
					}
				}
				isHighlighted = ((Operators.CompareString(strQuery, AH.A(113342), TextCompare: false) == 0 && base.IsLibraryContent) ? 1 : 0);
				goto IL_0097;
			}
		}
		goto IL_0096;
		IL_0096:
		isHighlighted = 1;
		goto IL_0097;
		IL_0097:
		((BaseItem)this).IsHighlighted = (byte)isHighlighted != 0;
	}
}
