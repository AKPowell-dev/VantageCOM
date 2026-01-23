using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class IndexedObject
{
	[CompilerGenerated]
	private object A;

	[CompilerGenerated]
	private Shape A;

	[CompilerGenerated]
	private object B;

	[CompilerGenerated]
	private bool A;

	[CompilerGenerated]
	private bool B;

	public object SlideOrLayout
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = RuntimeHelpers.GetObjectValue(value);
		}
	}

	public Shape Shape
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

	public object Child
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[CompilerGenerated]
		set
		{
			this.B = RuntimeHelpers.GetObjectValue(value);
		}
	}

	public bool IsMarker
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

	public bool AreRadarLabels
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public IndexedObject(object parent, Shape shp, object obj, bool isMarker = false, bool areRadarLabels = false)
	{
		SlideOrLayout = RuntimeHelpers.GetObjectValue(parent);
		Shape = shp;
		Child = RuntimeHelpers.GetObjectValue(obj);
		IsMarker = isMarker;
		AreRadarLabels = areRadarLabels;
	}
}
