using System.Collections.ObjectModel;
using System.Windows.Media;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Workbook.Merge;

public sealed class SourceWorkbook : BaseItem
{
	private new Microsoft.Office.Interop.Excel.Workbook A;

	private new string A;

	private new bool? A;

	private new bool A;

	private bool B;

	private new ImageSource A;

	private ImageSource B;

	private new Color A;

	private new double A;

	private new ObservableCollection<SourceSheet> A;

	public Microsoft.Office.Interop.Excel.Workbook Workbook
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			A(VH.A(175987));
		}
	}

	public string Path
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			A(VH.A(176004));
		}
	}

	public bool? IsChecked
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			if (value.HasValue)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (!value.Value)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						break;
					}
					TextColor = Colors.DarkGray;
					Opacity = 0.6;
					goto IL_0072;
				}
			}
			TextColor = Colors.Black;
			Opacity = 1.0;
			goto IL_0072;
			IL_0072:
			A(VH.A(90018));
		}
	}

	public bool IsSelected
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			A(VH.A(21693));
		}
	}

	public bool IsExpanded
	{
		get
		{
			return this.B;
		}
		set
		{
			this.B = value;
			A(VH.A(21595));
		}
	}

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

	public ImageSource FolderIcon
	{
		get
		{
			return B;
		}
		set
		{
			B = value;
			A(VH.A(176013));
		}
	}

	public Color TextColor
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			A(VH.A(90037));
		}
	}

	public double Opacity
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			A(VH.A(123854));
		}
	}

	public ObservableCollection<SourceSheet> Sheets
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
			A(VH.A(123841));
		}
	}

	public SourceWorkbook(Microsoft.Office.Interop.Excel.Workbook wb)
	{
		Workbook = wb;
		base.Name = wb.Name;
		string path;
		if (wb.Path.Length <= 0)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			path = wb.Path;
		}
		else
		{
			path = VH.A(115970) + wb.Path + VH.A(39904);
		}
		Path = path;
		IsChecked = true;
		IsExpanded = false;
		TextColor = Colors.Black;
		Opacity = 1.0;
		FileIcon = Forms.GetImageSource(J.ExcelSmall);
		FolderIcon = Forms.GetImageSource(J.FolderOpen);
		Sheets = new ObservableCollection<SourceSheet>();
	}
}
