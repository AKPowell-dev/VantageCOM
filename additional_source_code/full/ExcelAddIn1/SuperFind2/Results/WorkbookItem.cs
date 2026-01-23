using System.Collections.ObjectModel;
using A;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class WorkbookItem : BaseItem
{
	private ObservableCollection<SheetItem> m_A;

	private double m_A;

	public ObservableCollection<SheetItem> Sheets
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(123841));
		}
	}

	public double Opacity
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(123854));
		}
	}

	public WorkbookItem(Microsoft.Office.Interop.Excel.Workbook wb)
		: base(wb.Name, VH.A(123869))
	{
		base.Workbook = wb;
		Opacity = 1.0;
		((BaseItem)this).IndentLevel = 0;
		((BaseItem)this).IsExpanded = true;
		Sheets = new ObservableCollection<SheetItem>();
	}

	internal void A()
	{
		((BaseItem)this).Label = base.Workbook.Name;
	}
}
