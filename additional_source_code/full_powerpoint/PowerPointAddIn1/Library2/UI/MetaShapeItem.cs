using System.Xml;
using MacabacusMacros.Libraries.Pane.UI;

namespace PowerPointAddIn1.Library2.UI;

public class MetaShapeItem : ContentItem
{
	public MetaShapeItem(ContentGroup grp, XmlNode nd)
		: base(grp, nd)
	{
	}
}
