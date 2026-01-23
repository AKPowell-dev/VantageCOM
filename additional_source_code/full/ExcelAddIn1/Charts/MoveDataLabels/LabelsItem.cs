using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;

namespace ExcelAddIn1.Charts.MoveDataLabels;

public sealed class LabelsItem
{
	[CompilerGenerated]
	private string A;

	[CompilerGenerated]
	private Brush A;

	[CompilerGenerated]
	private Visibility A;

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

	public Brush Brush
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

	public Visibility ColorVisibility
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

	public LabelsItem(string strLabel, Brush br, Visibility vis)
	{
		Label = strLabel;
		Brush = br;
		ColorVisibility = vis;
	}
}
