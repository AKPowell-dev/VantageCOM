using System.Runtime.CompilerServices;
using System.Windows;

namespace PowerPointAddIn1.Pagination;

public sealed class SlidePair
{
	[CompilerGenerated]
	private BaseSlideItem A;

	[CompilerGenerated]
	private BaseSlideItem B;

	[CompilerGenerated]
	private int A;

	[CompilerGenerated]
	private Visibility A;

	public BaseSlideItem Left
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

	public BaseSlideItem Right
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

	public int PageNumber
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

	public Visibility SpiralVisibility
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

	public SlidePair(BaseSlideItem LeftSlide, BaseSlideItem RightSlide)
	{
		Left = LeftSlide;
		Right = RightSlide;
		SpiralVisibility = Visibility.Visible;
	}

	public SlidePair(BaseSlideItem LeftSlide, BaseSlideItem RightSlide, int intPageNumber)
	{
		Left = LeftSlide;
		Right = RightSlide;
		PageNumber = intPageNumber;
	}
}
