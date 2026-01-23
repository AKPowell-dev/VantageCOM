using System.Runtime.CompilerServices;
using System.Windows.Media;
using A;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.SuperFind2.Results;

public abstract class BaseItem : BaseItem
{
	[CompilerGenerated]
	private Microsoft.Office.Interop.Excel.Workbook A;

	private Brush A;

	private Brush B;

	internal Microsoft.Office.Interop.Excel.Workbook Workbook
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public Brush FontColor
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(124316));
		}
	}

	public Brush IconColor
	{
		get
		{
			return B;
		}
		set
		{
			B = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(124335));
		}
	}

	public BaseItem(string strLabel, string strData)
		: base(strLabel, strData)
	{
		FontColor = new SolidColorBrush(base.DEFAULT_FONT_COLOR);
		IconColor = FontColor;
	}

	public BaseItem(string strLabel, Geometry geo)
		: base(strLabel, geo)
	{
		FontColor = new SolidColorBrush(base.DEFAULT_FONT_COLOR);
		IconColor = FontColor;
	}
}
