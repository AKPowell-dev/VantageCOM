using System.Xml;
using MacabacusMacros.Libraries.Pane.UI;

namespace PowerPointAddIn1.Library2.UI;

public sealed class PitchlyItem : MetaShapeItem
{
	public PitchlyItem(ContentGroup grp, XmlNode nd)
		: base(grp, nd)
	{
	}
}
