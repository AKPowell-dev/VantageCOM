using System;
using System.Windows.Media;

namespace PowerPointAddIn1.Explorer;

public sealed class DummyItem : ContentItem
{
	public DummyItem()
		: base(null, "", new SolidColorBrush(System.Windows.Media.Colors.Black), Pane.CachedObjects.GeoChart)
	{
	}

	public override void Refresh()
	{
		throw new NotImplementedException();
	}

	public override void Delete()
	{
		throw new NotImplementedException();
	}

	public override void Search(string strQuery)
	{
	}
}
