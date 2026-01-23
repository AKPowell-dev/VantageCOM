using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Links;

public sealed class GlowBorder
{
	[CompilerGenerated]
	private int A;

	[CompilerGenerated]
	private float A;

	[CompilerGenerated]
	private float B;

	public int RBGColor
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

	public float Transparency
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

	public float Radius
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

	public GlowBorder(GlowFormat format)
	{
		GlowFormat glowFormat = format;
		Radius = glowFormat.Radius;
		RBGColor = glowFormat.Color.RGB;
		Transparency = glowFormat.Transparency;
		glowFormat = null;
	}
}
