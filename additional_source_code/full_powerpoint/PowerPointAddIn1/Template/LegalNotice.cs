using System.Runtime.CompilerServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.Template;

public sealed class LegalNotice
{
	[CompilerGenerated]
	private int A;

	[CompilerGenerated]
	private int B;

	[CompilerGenerated]
	private int C;

	[CompilerGenerated]
	private string A;

	public int Paragraphs
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

	public int Lines
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

	public int Length
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	public string Text
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

	public LegalNotice(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		TextRange2 textRange = shp.TextFrame2.TextRange;
		Paragraphs = textRange.get_Paragraphs(-1, -1).Count;
		Lines = textRange.get_Lines(-1, -1).Count;
		Length = textRange.Text.Length;
		Text = textRange.Text;
		textRange = null;
	}
}
