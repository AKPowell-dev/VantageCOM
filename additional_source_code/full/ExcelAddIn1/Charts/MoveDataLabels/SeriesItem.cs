using System;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts.MoveDataLabels;

public sealed class SeriesItem
{
	internal Series A;

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

	public SeriesItem(Series ser, int idx, Visibility vis)
	{
		string text = string.Empty;
		try
		{
			text = this.A.Name;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (text.Length != 0)
			{
				goto IL_007a;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		text = VH.A(56404) + idx;
		goto IL_007a;
		IL_007a:
		Label = text;
		this.A = ser;
		ColorVisibility = vis;
		try
		{
			XlChartType chartType = ser.ChartType;
			if (chartType <= XlChartType.xlRadar)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					break;
				}
				if (chartType == XlChartType.xlXYScatter)
				{
					goto IL_0164;
				}
				if (chartType != XlChartType.xlRadar)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						break;
					}
					goto IL_0197;
				}
			}
			else if (chartType != XlChartType.xl3DLine)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					break;
				}
				if (chartType != XlChartType.xlLine)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						break;
					}
					switch (chartType)
					{
					case XlChartType.xlLineStacked:
					case XlChartType.xlLineStacked100:
					case XlChartType.xlLineMarkers:
					case XlChartType.xlLineMarkersStacked:
					case XlChartType.xlLineMarkersStacked100:
						break;
					case XlChartType.xlXYScatterSmooth:
					case XlChartType.xlXYScatterSmoothNoMarkers:
					case XlChartType.xlXYScatterLines:
					case XlChartType.xlXYScatterLinesNoMarkers:
					case XlChartType.xlRadarMarkers:
						goto IL_0164;
					default:
						goto IL_0197;
					}
				}
			}
			Brush = ColorTile.A(ser.Format.Line);
			return;
			IL_0197:
			Brush = ColorTile.A(ser.Format.Fill);
			return;
			IL_0164:
			object brush;
			if (ser.MarkerBackgroundColor <= -1)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				brush = ColorTile.A();
			}
			else
			{
				brush = ColorTile.A(ser.MarkerBackgroundColor);
			}
			Brush = (Brush)brush;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Brush = ColorTile.A();
			ProjectData.ClearProjectError();
		}
	}
}
