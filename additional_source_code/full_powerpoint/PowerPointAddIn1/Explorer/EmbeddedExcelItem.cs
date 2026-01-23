using System.Windows.Forms;
using System.Windows.Media;
using A;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Links;

namespace PowerPointAddIn1.Explorer;

public sealed class EmbeddedExcelItem : ContentItem
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

	public EmbeddedExcelItem(SlideItem wsi, Shape shp, SolidColorBrush brush)
		: base(wsi, shp.Name, brush, Pane.CachedObjects.GeoExcel)
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
		if (MessageBox.Show(AH.A(113231), AH.A(5874), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
		{
			base.Shape.Delete();
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
			if (Operators.CompareString(strQuery, AH.A(113736), TextCompare: false) != 0)
			{
				if (Operators.CompareString(strQuery, AH.A(113331), TextCompare: false) == 0)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						break;
					}
					if (base.IsLinked)
					{
						goto IL_00a6;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						break;
					}
				}
				if (Operators.CompareString(strQuery, AH.A(113342), TextCompare: false) == 0)
				{
					while (true)
					{
						switch (3)
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
				goto IL_00a7;
			}
		}
		goto IL_00a6;
		IL_00a7:
		((BaseItem)this).IsHighlighted = (byte)isHighlighted != 0;
		return;
		IL_00a6:
		isHighlighted = 1;
		goto IL_00a7;
	}
}
