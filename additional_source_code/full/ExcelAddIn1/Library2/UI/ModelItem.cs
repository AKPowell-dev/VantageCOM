using System.Xml;
using MacabacusMacros.Libraries.Pane.UI;

namespace ExcelAddIn1.Library2.UI;

public sealed class ModelItem : ContentItem
{
	public ModelItem(ContentGroup grp, XmlNode nd)
		: base(grp, nd)
	{
	}

	public override void AsynchLoad()
	{
		((ContentItem)this).AsynchLoad();
		((ContentItem)this).SetSlidesLabel();
	}
}
