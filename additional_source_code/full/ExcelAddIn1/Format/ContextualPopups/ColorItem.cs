using System.Drawing;
using System.Windows;
using System.Windows.Media;
using System.Xml;
using A;
using Microsoft.Office.Core;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format.ContextualPopups;

public sealed class ColorItem
{
	private string m_A;

	private MsoFillType m_A;

	private System.Windows.Media.Color m_A;

	private System.Windows.Media.Color B;

	private SolidColorBrush m_A;

	private SolidColorBrush B;

	private Visibility m_A;

	public string Name
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public MsoFillType Type
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public System.Windows.Media.Color ForeColor
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public System.Windows.Media.Color BackColor
	{
		get
		{
			return this.B;
		}
		set
		{
			this.B = value;
		}
	}

	public SolidColorBrush ForeBrush
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public SolidColorBrush BackBrush
	{
		get
		{
			return B;
		}
		set
		{
			B = value;
		}
	}

	public Visibility CheckVisibility
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public int OleForeColor => A(ForeColor);

	public int OleBackColor => A(BackColor);

	public ColorItem(XmlNode nd, bool blnChecked)
	{
		string value = nd.Attributes[VH.A(144922)].Value;
		object obj = System.Windows.Media.ColorConverter.ConvertFromString(nd.Attributes[VH.A(144941)].Value);
		System.Windows.Media.Color foreColor;
		if (obj == null)
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
			foreColor = default(System.Windows.Media.Color);
		}
		else
		{
			foreColor = (System.Windows.Media.Color)obj;
		}
		ForeColor = foreColor;
		ForeBrush = new SolidColorBrush(ForeColor);
		if (value.Length > 0)
		{
			object obj2 = System.Windows.Media.ColorConverter.ConvertFromString(value);
			System.Windows.Media.Color backColor;
			if (obj2 == null)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					break;
				}
				backColor = default(System.Windows.Media.Color);
			}
			else
			{
				backColor = (System.Windows.Media.Color)obj2;
			}
			BackColor = backColor;
			BackBrush = new SolidColorBrush(BackColor);
		}
		if (!blnChecked)
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
			CheckVisibility = Visibility.Hidden;
		}
		else
		{
			CheckVisibility = Visibility.Visible;
		}
		Name = nd.Attributes[VH.A(67336)].Value;
		Type = (MsoFillType)Conversions.ToInteger(nd.Attributes[VH.A(144960)].Value);
	}

	private int A(System.Windows.Media.Color A)
	{
		return ColorTranslator.ToOle(System.Drawing.Color.FromArgb(A.R, A.G, A.B));
	}
}
