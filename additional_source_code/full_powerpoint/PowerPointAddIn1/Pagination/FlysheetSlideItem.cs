using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media.Imaging;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.Pagination;

public sealed class FlysheetSlideItem : RegularSlideItem
{
	[CompilerGenerated]
	private HorizontalAlignment A;

	[CompilerGenerated]
	private Thickness A;

	public HorizontalAlignment AdornerAlignment
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public Thickness AdornerPadding
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

	public FlysheetSlideItem(Slide sld, BitmapImage img)
		: base(sld, img)
	{
		AdornerAlignment = HorizontalAlignment.Left;
		AdornerPadding = new Thickness(2.0, 0.0, 3.0, 0.0);
	}
}
