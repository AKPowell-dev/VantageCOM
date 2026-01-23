using System.Xml;
using MacabacusMacros.Libraries.Pane.UI;

namespace ExcelAddIn1.Library2;

public sealed class TableItem : ContentItem
{
	private int A;

	public int SheetIndex
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
		}
	}

	public TableItem(ContentGroup grp, XmlNode nd)
		: base(grp, nd)
	{
	}

	public override void AsynchLoad()
	{
		((ContentItem)this).AsynchLoad();
		SheetIndex = ((ContentItem)this).GetSheetIndex();
	}
}
