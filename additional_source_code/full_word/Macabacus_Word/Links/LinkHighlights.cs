using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Links;

public sealed class LinkHighlights
{
	[CompilerGenerated]
	private Dictionary<Range, WdColorIndex> A;

	[CompilerGenerated]
	private Dictionary<Table, TableBorders> A;

	[CompilerGenerated]
	private Dictionary<Shape, ShapeBorder> A;

	[CompilerGenerated]
	private Dictionary<InlineShape, InlineShapeBorder> A;

	[CompilerGenerated]
	private Dictionary<Shape, GlowBorder> A;

	[CompilerGenerated]
	private Dictionary<InlineShape, GlowBorder> A;

	public Dictionary<Range, WdColorIndex> Text
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

	public Dictionary<Table, TableBorders> Tables
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

	public Dictionary<Shape, ShapeBorder> Shapes
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

	public Dictionary<InlineShape, InlineShapeBorder> InlineShapes
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

	public Dictionary<Shape, GlowBorder> GlowShapes
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

	public Dictionary<InlineShape, GlowBorder> InlineGlowShapes
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

	public LinkHighlights()
	{
		Text = new Dictionary<Range, WdColorIndex>();
		Shapes = new Dictionary<Shape, ShapeBorder>();
		InlineShapes = new Dictionary<InlineShape, InlineShapeBorder>();
		InlineGlowShapes = new Dictionary<InlineShape, GlowBorder>();
		GlowShapes = new Dictionary<Shape, GlowBorder>();
		Tables = new Dictionary<Table, TableBorders>();
	}
}
