using System.Runtime.CompilerServices;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using A;
using MacabacusMacros.Explorer;

namespace ExcelAddIn1.SuperFind2.Results;

public abstract class ExploreItem : ResultItem
{
	[CompilerGenerated]
	private BitmapSource A;

	private string A;

	[CompilerGenerated]
	private ContextMenu A;

	public BitmapSource PreviewImage
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

	public string Tooltip
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(118701));
		}
	}

	public ContextMenu ContextMenu
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

	public ExploreItem(WorksheetItem wsi, SolidColorBrush brush, Geometry geo, int index)
		: base(wsi, wsi.Worksheet, "", geo, index)
	{
		PreviewImage = null;
		base.IconColor = brush;
		if (index <= 50)
		{
			while (true)
			{
				switch (1)
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
			if (index >= 1)
			{
				return;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		base.UiIndex = 50;
	}

	public abstract void Refresh();

	public abstract void Delete();

	public abstract void Search(string strQuery);

	public void SearchOnInstantiate()
	{
	}
}
