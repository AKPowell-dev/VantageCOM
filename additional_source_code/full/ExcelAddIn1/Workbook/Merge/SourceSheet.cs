using System.Runtime.CompilerServices;
using System.Windows.Media;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Workbook.Merge;

public sealed class SourceSheet : BaseItem
{
	private new object A;

	private new int A;

	private new bool A;

	private bool B;

	private new ImageSource A;

	private new Color A;

	private new double A;

	private new SourceWorkbook A;

	public object Sheet
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = RuntimeHelpers.GetObjectValue(value);
			A(VH.A(175976));
		}
	}

	public int Index
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			A(VH.A(48135));
		}
	}

	public bool IsChecked
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			if (value)
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
				TextColor = Colors.Black;
				Opacity = 1.0;
			}
			else
			{
				TextColor = Colors.DarkGray;
				Opacity = 0.6;
			}
			A(VH.A(90018));
		}
	}

	public bool IsSelected
	{
		get
		{
			return B;
		}
		set
		{
			B = value;
			A(VH.A(21693));
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

	public SourceWorkbook Parent
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
			A(VH.A(8701));
		}
	}

	public SourceSheet(object sh, SourceWorkbook mw)
	{
		Sheet = RuntimeHelpers.GetObjectValue(sh);
		base.Name = Conversions.ToString(NewLateBinding.LateGet(sh, null, VH.A(19019), new object[0], null, null, null));
		Index = Conversions.ToInteger(NewLateBinding.LateGet(sh, null, VH.A(48135), new object[0], null, null, null));
		IsChecked = true;
		TextColor = Colors.Black;
		Opacity = 1.0;
		if (sh is Worksheet)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			SheetIcon = Forms.GetImageSource(J.Worksheet);
		}
		else
		{
			SheetIcon = Forms.GetImageSource(J.ChartInsert);
		}
		Parent = mw;
	}
}
