using System.Windows;
using System.Windows.Media;
using A;

namespace ExcelAddIn1.Workbook.Merge;

public sealed class DestinationSheet : BaseItem
{
	private new SourceSheet A;

	private new ImageSource A;

	private new Visibility A;

	public SourceSheet Source
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			A(VH.A(69016));
		}
	}

	public ImageSource SheetIcon
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			A(VH.A(175940));
		}
	}

	public Visibility Visibility
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
			A(VH.A(21785));
		}
	}

	public DestinationSheet(SourceSheet ss)
	{
		Source = ss;
		base.Name = ss.Name;
		SheetIcon = ss.SheetIcon;
		Visibility = Visibility.Visible;
	}
}
