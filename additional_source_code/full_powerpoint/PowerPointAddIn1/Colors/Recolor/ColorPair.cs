using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Media;
using A;
using MacabacusMacros;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Colors.Recolor;

public sealed class ColorPair : INotifyPropertyChanged
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	[CompilerGenerated]
	private bool m_A;

	private Color m_A;

	private Color? m_A;

	private SolidColorBrush m_A;

	private SolidColorBrush m_B;

	private SolidColorBrush C;

	private SolidColorBrush D;

	private SolidColorBrush E;

	private string m_A;

	private string m_B;

	private bool m_B;

	private Visibility m_A;

	private Visibility m_B;

	private Visibility C;

	private bool IsPaletteColor
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public Color OldColor
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(12406));
			OldFillBrush = new SolidColorBrush(value);
			OldBorderBrush = A(value);
			A();
			int warningVisible;
			if (!IsPaletteColor)
			{
				while (true)
				{
					switch (2)
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
				warningVisible = 0;
			}
			else
			{
				warningVisible = 2;
			}
			WarningVisible = (Visibility)warningVisible;
		}
	}

	public Color? NewColor
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(12423));
			if (value.HasValue)
			{
				while (true)
				{
					switch (3)
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
				NewFillBrush = new SolidColorBrush(value.Value);
				NewBorderBrush = A(value.Value);
				NewColorVisible = Visibility.Visible;
				WarningVisible = Visibility.Collapsed;
				object obj = ColorConverter.ConvertFromString(AH.A(12183));
				ArrowBrush = new SolidColorBrush((obj != null) ? ((Color)obj) : default(Color));
			}
			else
			{
				NewColorVisible = Visibility.Collapsed;
				int warningVisible;
				if (!IsPaletteColor)
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
					warningVisible = 0;
				}
				else
				{
					warningVisible = 2;
				}
				WarningVisible = (Visibility)warningVisible;
				object obj2 = ColorConverter.ConvertFromString(AH.A(12440));
				Color color;
				if (obj2 == null)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					color = default(Color);
				}
				else
				{
					color = (Color)obj2;
				}
				ArrowBrush = new SolidColorBrush(color);
			}
			B();
		}
	}

	public SolidColorBrush OldFillBrush
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(12455));
		}
	}

	public SolidColorBrush NewFillBrush
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(AH.A(12480));
		}
	}

	public SolidColorBrush OldBorderBrush
	{
		get
		{
			return this.C;
		}
		set
		{
			this.C = value;
			A(AH.A(12505));
		}
	}

	public SolidColorBrush NewBorderBrush
	{
		get
		{
			return D;
		}
		set
		{
			D = value;
			A(AH.A(12534));
		}
	}

	public SolidColorBrush ArrowBrush
	{
		get
		{
			return E;
		}
		set
		{
			E = value;
			A(AH.A(12563));
		}
	}

	public string OldLabel
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(12584));
		}
	}

	public string NewLabel
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(AH.A(12601));
		}
	}

	private bool IsLabelRgb
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A();
			B();
		}
	}

	public Visibility NewColorVisible
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(12618));
			int pickerButtonVisible;
			if (value != Visibility.Visible)
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
				pickerButtonVisible = 0;
			}
			else
			{
				pickerButtonVisible = 2;
			}
			PickerButtonVisible = (Visibility)pickerButtonVisible;
		}
	}

	public Visibility PickerButtonVisible
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(AH.A(12649));
		}
	}

	public Visibility WarningVisible
	{
		get
		{
			return C;
		}
		set
		{
			C = value;
			A(AH.A(12688));
		}
	}

	public event PropertyChangedEventHandler PropertyChanged
	{
		[CompilerGenerated]
		add
		{
			PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Combine(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.m_A, value2, propertyChangedEventHandler2);
			}
			while ((object)propertyChangedEventHandler != propertyChangedEventHandler2);
		}
		[CompilerGenerated]
		remove
		{
			PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Remove(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.m_A, value2, propertyChangedEventHandler2);
			}
			while ((object)propertyChangedEventHandler != propertyChangedEventHandler2);
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				return;
			}
		}
	}

	public ColorPair(Color clr, bool blnRgb)
	{
		this.m_B = Visibility.Visible;
		C = Visibility.Collapsed;
		IsPaletteColor = false;
		foreach (PaletteColor item in clsColors.ColorPalette)
		{
			if (Operators.CompareString(item.RGB, Conversions.ToString(clr.R) + AH.A(12717) + Conversions.ToString(clr.G) + AH.A(12717) + Conversions.ToString(clr.B), TextCompare: false) == 0)
			{
				IsPaletteColor = true;
				break;
			}
		}
		OldColor = clr;
		NewColor = null;
		GenerateLabels(blnRgb);
	}

	private void A(string A)
	{
		this.m_A?.Invoke(this, new PropertyChangedEventArgs(A));
	}

	public void GenerateLabels(bool blnRgb)
	{
		IsLabelRgb = blnRgb;
	}

	public void Reset()
	{
		NewColor = null;
	}

	private SolidColorBrush A(Color A)
	{
		if (!(A == System.Windows.Media.Colors.White))
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return new SolidColorBrush(A);
				}
			}
		}
		object obj = ColorConverter.ConvertFromString(AH.A(12440));
		Color color;
		if (obj == null)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			color = default(Color);
		}
		else
		{
			color = (Color)obj;
		}
		return new SolidColorBrush(color);
	}

	private void A()
	{
		OldLabel = A(OldColor, IsLabelRgb);
	}

	private void B()
	{
		if (NewColor.HasValue)
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
					NewLabel = A(NewColor.Value, IsLabelRgb);
					return;
				}
			}
		}
		NewLabel = "";
	}

	private string A(Color A, bool B)
	{
		if (B)
		{
			return Conversions.ToString(A.R) + AH.A(7894) + Conversions.ToString(A.G) + AH.A(7894) + Conversions.ToString(A.B);
		}
		string text = Strings.Right(A.ToString(), 6);
		return text.Substring(0, 2) + AH.A(7894) + text.Substring(2, 2) + AH.A(7894) + text.Substring(4, 2);
	}
}
