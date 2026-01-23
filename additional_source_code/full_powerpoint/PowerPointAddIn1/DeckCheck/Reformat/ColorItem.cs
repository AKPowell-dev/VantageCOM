using System.Collections.Generic;
using System.Drawing;
using System.Windows;
using System.Windows.Media;
using A;
using MacabacusMacros;
using MacabacusMacros.Proofing.UI.Reformat;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public abstract class ColorItem : BaseItem
{
	private System.Drawing.Color A;

	private SolidColorBrush A;

	private SolidColorBrush B;

	public System.Drawing.Color Color
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(13587));
		}
	}

	public SolidColorBrush FillColor
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(49852));
		}
	}

	public SolidColorBrush BorderColor
	{
		get
		{
			return B;
		}
		set
		{
			B = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(49871));
		}
	}

	public ColorItem(int intColor, List<IndexedObject> listObjects, int intIndex, DataTemplate template, DataTemplate navItemTemplate, string strHeader, int intTotal)
		: base(intTotal, intIndex, listObjects, template, navItemTemplate, strHeader)
	{
		Color = ColorTranslator.FromOle(intColor);
		System.Windows.Media.Color color = System.Windows.Media.Color.FromRgb(Color.R, Color.G, Color.B);
		FillColor = new SolidColorBrush(color);
		if (color != System.Windows.Media.Colors.White)
		{
			BorderColor = FillColor;
		}
		else
		{
			BorderColor = new SolidColorBrush(System.Windows.Media.Colors.LightGray);
		}
		((BaseItem)this).Label = AH.A(49894) + clsColors.Color2RGB(Color) + AH.A(49903) + color.ToString().Replace(AH.A(49910), AH.A(49917));
	}

	public abstract void Reformat(int intNewColor, int intOldColor, ref List<string> listErrors);
}
