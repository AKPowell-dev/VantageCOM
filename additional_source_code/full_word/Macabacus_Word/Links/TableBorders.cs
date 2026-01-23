using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Links;

public sealed class TableBorders
{
	[CompilerGenerated]
	private TableBorder A;

	[CompilerGenerated]
	private TableBorder B;

	[CompilerGenerated]
	private TableBorder C;

	[CompilerGenerated]
	private TableBorder D;

	public TableBorder Bottom
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

	public TableBorder Left
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

	public TableBorder Right
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

	public TableBorder Top
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[CompilerGenerated]
		set
		{
			D = value;
		}
	}

	public TableBorders(Table tbl)
	{
		Table table = tbl;
		Bottom = new TableBorder(table.Borders[WdBorderType.wdBorderBottom]);
		Left = new TableBorder(table.Borders[WdBorderType.wdBorderLeft]);
		Right = new TableBorder(table.Borders[WdBorderType.wdBorderRight]);
		Top = new TableBorder(table.Borders[WdBorderType.wdBorderTop]);
		table = null;
	}
}
