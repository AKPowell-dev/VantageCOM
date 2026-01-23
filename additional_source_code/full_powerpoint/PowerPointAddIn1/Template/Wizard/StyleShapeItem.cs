using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.Template.Wizard;

public sealed class StyleShapeItem
{
	[CompilerGenerated]
	private string A;

	[CompilerGenerated]
	private Shape A;

	public string Name
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

	public Shape Shape
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

	public StyleShapeItem(Shape shp)
	{
		Shape = shp;
		Name = shp.Name;
	}
}
