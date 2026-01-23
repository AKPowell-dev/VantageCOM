using System.Runtime.CompilerServices;
using System.Windows.Media;
using System.Xml;
using A;

namespace ExcelAddIn1.Charts;

public sealed class StandardSizeItem
{
	[CompilerGenerated]
	private string A;

	[CompilerGenerated]
	private Geometry A;

	public string Label
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

	public Geometry Icon
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

	public StandardSizeItem()
	{
		Label = VH.A(65507);
		Icon = Geometry.Parse(VH.A(65520));
	}

	public StandardSizeItem(XmlNode nd)
	{
		Label = nd.Attributes[VH.A(67336)].Value;
		Icon = Geometry.Parse(VH.A(67345));
	}
}
