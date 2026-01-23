using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using A;

namespace Macabacus_Word.Links;

public sealed class SortAdorner : Adorner
{
	private static Geometry A = Geometry.Parse(XC.A(17927));

	private static Geometry B = Geometry.Parse(XC.A(17970));

	[CompilerGenerated]
	private ListSortDirection A;

	public ListSortDirection Direction
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

	public SortAdorner(UIElement element, ListSortDirection dir)
		: base(element)
	{
		Direction = dir;
	}

	protected override void OnRender(DrawingContext drawingContext)
	{
		base.OnRender(drawingContext);
		if (base.AdornedElement.RenderSize.Width < 20.0)
		{
			return;
		}
		TranslateTransform transform = new TranslateTransform(base.AdornedElement.RenderSize.Width - 15.0, (base.AdornedElement.RenderSize.Height - 5.0) / 2.0);
		drawingContext.PushTransform(transform);
		Geometry geometry = SortAdorner.A;
		if (Direction == ListSortDirection.Descending)
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
			geometry = B;
		}
		drawingContext.DrawGeometry(Brushes.Gray, null, geometry);
		drawingContext.Pop();
	}
}
