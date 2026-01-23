using System.Runtime.CompilerServices;
using System.Windows.Media;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.TraceDialogs.Precedents;

public class BaseItem : TraceItem
{
	private string A;

	private string B;

	[CompilerGenerated]
	private int A;

	[CompilerGenerated]
	private int B;

	[CompilerGenerated]
	private bool A;

	private bool B;

	private SolidColorBrush A;

	public string Value
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			NotifyPropertyChanged(VH.A(41636));
		}
	}

	public string Info
	{
		get
		{
			return this.B;
		}
		set
		{
			this.B = value;
			NotifyPropertyChanged(VH.A(41647));
		}
	}

	public int SelectionIndex
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

	public int SelectionLength
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	public int SelectionEnd => checked(SelectionIndex - 1 + SelectionLength);

	public bool IsName
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

	public bool IsError
	{
		get
		{
			return B;
		}
		set
		{
			B = value;
			NotifyPropertyChanged(VH.A(44198));
			if (value)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						FontBrush = Brushes.Firebrick;
						return;
					}
				}
			}
			object obj = ColorConverter.ConvertFromString(VH.A(44213));
			FontBrush = new SolidColorBrush((obj != null) ? ((Color)obj) : default(Color));
		}
	}

	public SolidColorBrush FontBrush
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
			NotifyPropertyChanged(VH.A(44228));
		}
	}

	public BaseItem(BaseItem p, Range rng, int intLevel, string strIcon)
		: base(p, rng, intLevel, strIcon)
	{
		if (rng != null)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					IsError = Dialog.IsFormulaError(rng);
					return;
				}
			}
		}
		IsError = false;
	}

	public BaseItem FirstRangeParent()
	{
		int num = 0;
		BaseItem baseItem = this;
		do
		{
			baseItem = (BaseItem)baseItem.Parent;
			if (baseItem == null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return baseItem;
					}
				}
			}
			if (baseItem.Range != null)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						return baseItem;
					}
				}
			}
			num = checked(num + 1);
		}
		while (num != 2000);
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			return null;
		}
	}
}
