using System.Xml;
using MacabacusMacros.Libraries.Pane.UI;

namespace Macabacus_Word.Library2.UI;

public sealed class DocumentItem : ContentItem
{
	public DocumentItem(ContentGroup grp, XmlNode nd)
		: base(grp, nd)
	{
	}

	public override void AsynchLoad()
	{
		((ContentItem)this).AsynchLoad();
		((ContentItem)this).SetSlidesLabel();
	}
}
