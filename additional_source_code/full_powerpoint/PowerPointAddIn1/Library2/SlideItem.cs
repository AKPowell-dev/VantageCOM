using System.Runtime.CompilerServices;
using System.Xml;
using MacabacusMacros.Libraries.Pane.UI;

namespace PowerPointAddIn1.Library2;

public sealed class SlideItem : ContentItem
{
	[CompilerGenerated]
	private string A;

	public string LinkId
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

	public SlideItem(ContentGroup grp, XmlNode nd)
		: base(grp, nd)
	{
	}

	public override void AsynchLoad()
	{
		((ContentItem)this).AsynchLoad();
		((ContentItem)this).SetSlidesLabel();
		LinkId = ((ContentItem)this).GetLinkId();
	}
}
