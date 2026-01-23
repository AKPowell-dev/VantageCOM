using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using A;

namespace PowerPointAddIn1.Links;

public sealed class SortAdorner : Adorner
{
	private static Geometry A = Geometry.Parse(AH.A(98683));

	private static Geometry B = Geometry.Parse(AH.A(98726));

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
			TranslateTransform transform = new TranslateTransform(base.AdornedElement.RenderSize.Width - 15.0, (base.AdornedElement.RenderSize.Height - 5.0) / 2.0);
			drawingContext.PushTransform(transform);
			Geometry geometry = SortAdorner.A;
			if (Direction == ListSortDirection.Descending)
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
				geometry = B;
			}
			drawingContext.DrawGeometry(Brushes.Gray, null, geometry);
			drawingContext.Pop();
			return;
		}
	}
}
