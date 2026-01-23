using System.Xml;
using MacabacusMacros.Libraries.Pane.UI;

namespace PowerPointAddIn1.Library2.UI;

public sealed class DeckItem : ContentItem
{
	public DeckItem(ContentGroup grp, XmlNode nd)
		: base(grp, nd)
	{
	}

	public override void AsynchLoad()
	{
		((ContentItem)this).AsynchLoad();
		((ContentItem)this).SetSlidesLabel();
	}
}
