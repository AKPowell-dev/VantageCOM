using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Media;
using A;
using MacabacusMacros.UI;

namespace ExcelAddIn1.Workbook.Merge;

public sealed class DestinationWorkbook : BaseItem
{
	private new ImageSource A;

	private new ObservableCollection<DestinationSheet> A;

	private new Visibility A;

	public ImageSource FileIcon
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			A(VH.A(175959));
		}
	}

	public ObservableCollection<DestinationSheet> Sheets
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			A(VH.A(123841));
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

	public DestinationWorkbook()
	{
		FileIcon = Forms.GetImageSource(J.ExcelSmall);
		Sheets = new ObservableCollection<DestinationSheet>();
		Visibility = Visibility.Visible;
	}
}
