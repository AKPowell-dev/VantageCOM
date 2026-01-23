using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.Pagination;

public class BaseSlideItem
{
	[CompilerGenerated]
	private Slide A;

	public Slide Slide
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
}
