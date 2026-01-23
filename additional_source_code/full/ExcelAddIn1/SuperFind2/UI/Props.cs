using System.Runtime.CompilerServices;
using System.Windows.Controls;
using ExcelAddIn1.SuperFind2.Queries;

namespace ExcelAddIn1.SuperFind2.UI;

public sealed class Props
{
	[CompilerGenerated]
	private static BaseQuery A;

	[CompilerGenerated]
	private static ControlTemplate A;

	[CompilerGenerated]
	private static ControlTemplate B;

	[CompilerGenerated]
	private static IconCache A;

	internal static BaseQuery SearchForm
	{
		[CompilerGenerated]
		get
		{
			return Props.A;
		}
		[CompilerGenerated]
		set
		{
			Props.A = value;
		}
	}

	internal static ControlTemplate SingleCellIcon
	{
		[CompilerGenerated]
		get
		{
			return Props.A;
		}
		[CompilerGenerated]
		set
		{
			Props.A = value;
		}
	}

	internal static ControlTemplate MultiCellIcon
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

	internal static IconCache Icons
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
