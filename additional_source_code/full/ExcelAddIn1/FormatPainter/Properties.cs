using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FormatPainter;

public sealed class Properties
{
	public struct ChartObjectProperties
	{
		public MsoTriState LockAspectRatio;

		public double Top;

		public double Left;

		public double Width;

		public double Height;
	}

	public struct ChartProperties
	{
		public bool HasDataTable;

		public bool HasLegend;

		public bool HasTitle;

		public bool HasPrimaryValueAxis;

		public bool HasPrimaryCategoryAxis;

		public bool HasSecondaryValueAxis;

		public bool HasSecondaryCategoryAxis;

		public int ChartStyle;
	}

	public struct ChartAreaProperties
	{
		public FillProperties Fill;

		public LineProperties Border;
	}

	public struct PlotAreaProperties
	{
		public double InsideTop;

		public double InsideLeft;

		public double InsideHeight;

		public double InsideWidth;

		public FillProperties Fill;

		public LineProperties Border;
	}

	public struct SeriesProperties
	{
		public Dictionary<int, XlChartType> ChartType;

		public Dictionary<int, FillProperties> Fill;

		public Dictionary<int, LineProperties> Border;

		public Dictionary<int, MarkerProperties> Markers;

		public Dictionary<int, int> GapWidth;

		public Dictionary<int, int> Overlap;

		public Dictionary<int, int> FirstSliceAngle;

		public Dictionary<int, int> Explosion;

		public Dictionary<int, ErrorBarsProperties> ErrorBars;

		public Dictionary<int, UpDownBars> UpBars;

		public Dictionary<int, UpDownBars> DownBars;

		public Dictionary<int, DataLabelProperties> DataLabels;
	}

	public struct LegendProperties
	{
		public FontProperties Font;

		public FillProperties Fill;

		public LineProperties Border;

		public bool IncludeInLayout;

		public XlLegendPosition Position;

		public double Top;

		public double Left;
	}

	public struct TitleProperties
	{
		public FontProperties Font;

		public LineProperties Border;

		public FillProperties Fill;

		public bool IncludeInLayout;

		public XlChartElementPosition Position;

		public double Top;

		public double Left;
	}

	public struct DataTableProperties
	{
		public FontProperties Font;

		public LineProperties Border;

		public FillProperties Fill;

		public bool ShowLegendKey;

		public bool HasBorderHorizontal;

		public bool HasBorderOutline;

		public bool HasBorderVertical;
	}

	public struct PrimaryValueAxisProperties
	{
		public AxisProperties Axis;
	}

	public struct PrimaryCategoryAxisProperties
	{
		public AxisProperties Axis;
	}

	public struct SecondaryValueAxisProperties
	{
		public AxisProperties Axis;
	}

	public struct SecondaryCategoryAxisProperties
	{
		public AxisProperties Axis;
	}

	public struct LineProperties
	{
		public int ForeColor;

		public MsoLineStyle Style;

		public MsoLineDashStyle DashStyle;

		public float Weight;

		public float Transparency;

		public MsoTriState Visible;
	}

	public struct BorderProperties
	{
		public int Color;

		public XlLineStyle LineStyle;

		public float Weight;
	}

	public struct FillProperties
	{
		public int BackColor;

		public int ForeColor;

		public MsoThemeColorIndex ObjectThemeColor;

		public MsoPatternType Pattern;

		public float Transparency;

		public MsoFillType Type;

		public MsoTriState Visible;

		public float GradientAngle;

		public MsoGradientColorType GradientColorType;

		public float GradientDegree;

		public List<GradientStopProperties> GradientStops;

		public MsoGradientStyle GradientStyle;

		public int GradientVariant;

		public MsoPresetGradientType PresetGradientType;
	}

	public struct GradientStopProperties
	{
		public int RGB;

		public int SchemeColor;

		public float Brightness;

		public float Position;

		public float Transparency;

		public MsoColorType Type;

		public MsoThemeColorIndex ObjectThemeColor;
	}

	public struct FontProperties
	{
		public float Size;

		public int ForeColor;

		public int BackColor;

		public string Name;

		public DecorationProperties Decoration;
	}

	public struct DecorationProperties
	{
		public MsoTriState Bold;

		public MsoTriState Italic;

		public MsoTextUnderlineType UnderlineStyle;
	}

	public struct AxisProperties
	{
		public float MinimumScale;

		public bool MinimumScaleIsAuto;

		public double MaximumScale;

		public bool MaximumScaleIsAuto;

		public double MajorUnit;

		public bool MajorUnitIsAuto;

		public XlTimeUnit MajorUnitScale;

		public double MinorUnit;

		public bool MinorUnitIsAuto;

		public XlTimeUnit MinorUnitScale;

		public XlAxisCrosses Crosses;

		public double CrossesAt;

		public XlScaleType ScaleType;

		public bool ReversePlotOrder;

		public bool HasDisplayUnitLabel;

		public bool HasMajorGridlines;

		public bool HasMinorGridlines;

		public GridlinesProperties MajorGridlines;

		public GridlinesProperties MinorGridlines;

		public bool AxisBetweenCategories;

		public XlDisplayUnit DisplayUnit;

		public BorderProperties Border;

		public XlTickMark MajorTickMark;

		public XlTickMark MinorTickMark;

		public XlTickLabelPosition TickLabelPosition;

		public int TickLabelSpacing;

		public bool TickLabelSpacingIsAuto;

		public int TickMarkSpacing;

		public TickLabelsProperties TickLabels;

		public bool HasTitle;

		public AxisTitleProperties AxisTitle;
	}

	public struct GridlinesProperties
	{
		public LineProperties Format;
	}

	public struct TickLabelsProperties
	{
		public FontProperties Format;

		public bool MultiLevel;

		public int Offset;

		public int Orientation;

		public bool NumberFormatLinked;

		public string NumberFormat;

		public int Alignment;
	}

	public struct AxisTitleProperties
	{
		public FontProperties Font;

		public LineProperties Border;

		public FillProperties Fill;

		public bool IncludeInLayout;

		public object HorizontalAlignment;

		public object VerticalAlignment;

		public object Orientation;
	}

	public struct ErrorBarsProperties
	{
		public bool HasErrorBars;

		public LineProperties Line;

		public XlEndStyleCap EndStyle;
	}

	public struct UpDownBars
	{
		public bool HasUpDownBars;

		public LineProperties Border;

		public FillProperties Fill;
	}

	public struct DataLabelProperties
	{
		public bool HasDataLabels;

		public bool HasLeaderLines;

		public FontProperties Font;

		public FillProperties Fill;

		public LineProperties Border;

		public LineProperties Line;

		public XlDataLabelPosition Position;

		public int Orientation;

		public object HorizontalAlignment;

		public object VerticalAlignment;

		public bool AutoText;

		public string NumberFormat;

		public bool NumberFormatLinked;

		public bool ShowLegendKey;
	}

	public struct MarkerProperties
	{
		public int MarkerSize;

		public XlMarkerStyle MarkerStyle;

		public int MarkerBackgroundColor;

		public int MarkerForegroundColor;
	}

	private ChartObjectProperties m_A;

	private ChartProperties m_A;

	private ChartAreaProperties m_A;

	private PlotAreaProperties m_A;

	private SeriesProperties m_A;

	private LegendProperties m_A;

	private TitleProperties m_A;

	private DataTableProperties m_A;

	private AxisProperties m_A;

	private AxisProperties B;

	private AxisProperties C;

	private AxisProperties D;

	private BitmapSource m_A;

	public ChartObjectProperties ChartObject
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public ChartProperties Chart
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public ChartAreaProperties ChartArea
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public PlotAreaProperties PlotArea
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public SeriesProperties Series
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public LegendProperties Legend
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public TitleProperties Title
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public DataTableProperties DataTable
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public AxisProperties PrimaryValueAxis
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public AxisProperties PrimaryCategoryAxis
	{
		get
		{
			return B;
		}
		set
		{
			B = value;
		}
	}

	public AxisProperties SecondaryValueAxis
	{
		get
		{
			return C;
		}
		set
		{
			C = value;
		}
	}

	public AxisProperties SecondaryCategoryAxis
	{
		get
		{
			return D;
		}
		set
		{
			D = value;
		}
	}

	public BitmapSource SourceImage
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public Properties(Chart cht)
	{
		checked
		{
			try
			{
				ChartObjectProperties chartObject = default(ChartObjectProperties);
				ChartObject chartObject2 = (ChartObject)cht.Parent;
				chartObject.Height = chartObject2.Height;
				chartObject.Width = chartObject2.Width;
				chartObject.Top = chartObject2.Top;
				chartObject.Left = chartObject2.Left;
				chartObject.LockAspectRatio = chartObject2.ShapeRange.LockAspectRatio;
				chartObject2 = null;
				ChartObject = chartObject;
				ChartProperties chart = default(ChartProperties);
				Chart chart2 = cht;
				chart.HasDataTable = chart2.HasDataTable;
				chart.HasLegend = chart2.HasLegend;
				chart.HasTitle = chart2.HasTitle;
				chart.HasPrimaryValueAxis = Conversions.ToBoolean(((_Chart)chart2).get_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlValue, (object)XlAxisGroup.xlPrimary));
				chart.HasPrimaryCategoryAxis = Conversions.ToBoolean(((_Chart)chart2).get_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, (object)XlAxisGroup.xlPrimary));
				chart.HasSecondaryValueAxis = Conversions.ToBoolean(((_Chart)chart2).get_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlValue, (object)XlAxisGroup.xlSecondary));
				chart.HasSecondaryCategoryAxis = Conversions.ToBoolean(((_Chart)chart2).get_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, (object)XlAxisGroup.xlSecondary));
				chart.ChartStyle = Conversions.ToInteger(chart2.ChartStyle);
				chart2 = null;
				Chart = chart;
				ChartAreaProperties chartArea = default(ChartAreaProperties);
				ChartFormat format = cht.ChartArea.Format;
				FillProperties fillProperties = GetFillProperties(format.Fill);
				LineProperties borderProperties = GetBorderProperties(format.Line);
				_ = null;
				chartArea.Fill = fillProperties;
				chartArea.Border = borderProperties;
				ChartArea = chartArea;
				PlotAreaProperties plotArea = default(PlotAreaProperties);
				PlotArea plotArea2 = cht.PlotArea;
				plotArea.InsideHeight = plotArea2.InsideHeight;
				plotArea.InsideWidth = plotArea2.InsideWidth;
				plotArea.InsideTop = plotArea2.InsideTop;
				plotArea.InsideLeft = plotArea2.InsideLeft;
				fillProperties = GetFillProperties(plotArea2.Format.Fill);
				borderProperties = GetBorderProperties(plotArea2.Format.Line);
				plotArea2 = null;
				plotArea.Fill = fillProperties;
				plotArea.Border = borderProperties;
				PlotArea = plotArea;
				SeriesProperties series = default(SeriesProperties);
				Dictionary<int, XlChartType> dictionary = new Dictionary<int, XlChartType>();
				Dictionary<int, FillProperties> dictionary2 = new Dictionary<int, FillProperties>();
				Dictionary<int, LineProperties> dictionary3 = new Dictionary<int, LineProperties>();
				Dictionary<int, MarkerProperties> dictionary4 = new Dictionary<int, MarkerProperties>();
				Dictionary<int, int> dictionary5 = new Dictionary<int, int>();
				Dictionary<int, int> dictionary6 = new Dictionary<int, int>();
				Dictionary<int, int> dictionary7 = new Dictionary<int, int>();
				Dictionary<int, int> dictionary8 = new Dictionary<int, int>();
				Dictionary<int, ErrorBarsProperties> dictionary9 = new Dictionary<int, ErrorBarsProperties>();
				Dictionary<int, DataLabelProperties> dictionary10 = new Dictionary<int, DataLabelProperties>();
				Dictionary<int, UpDownBars> dictionary11 = new Dictionary<int, UpDownBars>();
				Dictionary<int, UpDownBars> dictionary12 = new Dictionary<int, UpDownBars>();
				try
				{
					FullSeriesCollection fullSeriesCollection = (FullSeriesCollection)cht.FullSeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value));
					int count = fullSeriesCollection.Count;
					for (int i = 1; i <= count; i++)
					{
						Series series2 = fullSeriesCollection.Item(i);
						dictionary.Add(i, series2.ChartType);
						dictionary2.Add(i, GetFillProperties(series2.Format.Fill));
						dictionary3.Add(i, GetBorderProperties(series2.Format.Line));
						try
						{
							MarkerProperties value = new MarkerProperties
							{
								MarkerStyle = series2.MarkerStyle
							};
							if (series2.MarkerStyle != XlMarkerStyle.xlMarkerStyleNone)
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
								value.MarkerSize = series2.MarkerSize;
								value.MarkerForegroundColor = series2.MarkerForegroundColor;
								value.MarkerBackgroundColor = series2.MarkerBackgroundColor;
							}
							dictionary4.Add(i, value);
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						try
						{
							if (A(series2.ChartType))
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									dictionary8.Add(i, series2.Explosion);
									break;
								}
							}
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
						}
						try
						{
							ErrorBarsProperties value2 = new ErrorBarsProperties
							{
								HasErrorBars = series2.HasErrorBars
							};
							if (series2.HasErrorBars)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									break;
								}
								value2.EndStyle = series2.ErrorBars.EndStyle;
								value2.Line = GetBorderProperties(series2.ErrorBars.Format.Line);
							}
							dictionary9.Add(i, value2);
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
							ProjectData.ClearProjectError();
						}
						try
						{
							DataLabelProperties value3 = new DataLabelProperties
							{
								HasDataLabels = series2.HasDataLabels,
								HasLeaderLines = series2.HasLeaderLines
							};
							if (series2.HasDataLabels)
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
								DataLabels dataLabels = (DataLabels)series2.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
								value3.Font = GetFontProperties(dataLabels.Font);
								value3.Fill = GetFillProperties(dataLabels.Format.Fill);
								value3.Border = GetBorderProperties(dataLabels.Format.Line);
								value3.Position = dataLabels.Position;
								value3.Orientation = Conversions.ToInteger(dataLabels.Orientation);
								value3.HorizontalAlignment = RuntimeHelpers.GetObjectValue(dataLabels.HorizontalAlignment);
								value3.VerticalAlignment = RuntimeHelpers.GetObjectValue(dataLabels.VerticalAlignment);
								value3.AutoText = dataLabels.AutoText;
								value3.NumberFormat = dataLabels.NumberFormat;
								value3.NumberFormatLinked = dataLabels.NumberFormatLinked;
								value3.ShowLegendKey = dataLabels.ShowLegendKey;
								dataLabels = null;
							}
							if (series2.HasLeaderLines)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									break;
								}
								value3.Line = GetBorderProperties(series2.LeaderLines.Format.Line);
							}
							dictionary10.Add(i, value3);
						}
						catch (Exception ex7)
						{
							ProjectData.SetProjectError(ex7);
							Exception ex8 = ex7;
							ProjectData.ClearProjectError();
						}
						series2 = null;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						fullSeriesCollection = null;
						break;
					}
				}
				catch (Exception ex9)
				{
					ProjectData.SetProjectError(ex9);
					Exception ex10 = ex9;
					ProjectData.ClearProjectError();
				}
				try
				{
					ChartGroups chartGroups = (ChartGroups)cht.ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value));
					int count2 = chartGroups.Count;
					for (int j = 1; j <= count2; j++)
					{
						ChartGroup chartGroup = chartGroups.Item(j);
						try
						{
							dictionary5.Add(j, chartGroup.GapWidth);
							dictionary6.Add(j, chartGroup.Overlap);
						}
						catch (Exception ex11)
						{
							ProjectData.SetProjectError(ex11);
							Exception ex12 = ex11;
							ProjectData.ClearProjectError();
						}
						try
						{
							if (A(cht.ChartType))
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									dictionary7.Add(j, chartGroup.FirstSliceAngle);
									break;
								}
							}
						}
						catch (Exception ex13)
						{
							ProjectData.SetProjectError(ex13);
							Exception ex14 = ex13;
							ProjectData.ClearProjectError();
						}
						UpDownBars value4 = default(UpDownBars);
						UpDownBars upDownBars = default(UpDownBars);
						try
						{
							value4.HasUpDownBars = chartGroup.HasUpDownBars;
							upDownBars.HasUpDownBars = chartGroup.HasUpDownBars;
							if (chartGroup.HasUpDownBars)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									value4.Border = GetBorderProperties(chartGroup.UpBars.Format.Line);
									value4.Fill = GetFillProperties(chartGroup.UpBars.Format.Fill);
									upDownBars.Border = GetBorderProperties(chartGroup.DownBars.Format.Line);
									upDownBars.Fill = GetFillProperties(chartGroup.DownBars.Format.Fill);
									break;
								}
							}
						}
						catch (Exception ex15)
						{
							ProjectData.SetProjectError(ex15);
							Exception ex16 = ex15;
							ProjectData.ClearProjectError();
						}
						dictionary11.Add(j, value4);
						dictionary12.Add(j, value4);
						chartGroup = null;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						chartGroups = null;
						break;
					}
				}
				catch (Exception ex17)
				{
					ProjectData.SetProjectError(ex17);
					Exception ex18 = ex17;
					ProjectData.ClearProjectError();
				}
				series.ChartType = dictionary;
				series.Fill = dictionary2;
				series.Border = dictionary3;
				series.Markers = dictionary4;
				series.GapWidth = dictionary5;
				series.Overlap = dictionary6;
				series.FirstSliceAngle = dictionary7;
				series.Explosion = dictionary8;
				series.ErrorBars = dictionary9;
				series.UpBars = dictionary11;
				series.DownBars = dictionary12;
				series.DataLabels = dictionary10;
				dictionary = null;
				dictionary2 = null;
				dictionary3 = null;
				dictionary4 = null;
				dictionary5 = null;
				dictionary6 = null;
				dictionary7 = null;
				dictionary8 = null;
				dictionary9 = null;
				dictionary11 = null;
				dictionary12 = null;
				dictionary10 = null;
				Series = series;
				LegendProperties legend = default(LegendProperties);
				if (cht.HasLegend)
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
					Legend legend2 = cht.Legend;
					legend.Position = legend2.Position;
					legend.IncludeInLayout = legend2.IncludeInLayout;
					legend.Top = legend2.Top;
					legend.Left = legend2.Left;
					FontProperties fontProperties = GetFontProperties(legend2.Font);
					fillProperties = GetFillProperties(legend2.Format.Fill);
					borderProperties = GetBorderProperties(legend2.Format.Line);
					legend2 = null;
					legend.Font = fontProperties;
					legend.Fill = fillProperties;
					legend.Border = borderProperties;
				}
				Legend = legend;
				TitleProperties title = default(TitleProperties);
				if (cht.HasTitle)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
					ChartTitle chartTitle = cht.ChartTitle;
					FontProperties fontProperties;
					try
					{
						fontProperties = GetFontProperties(chartTitle.Format.TextFrame2.TextRange.Font);
					}
					catch (Exception ex19)
					{
						ProjectData.SetProjectError(ex19);
						Exception ex20 = ex19;
						fontProperties = GetFontProperties(chartTitle.Font);
						ProjectData.ClearProjectError();
					}
					fillProperties = GetFillProperties(chartTitle.Format.Fill);
					borderProperties = GetBorderProperties(chartTitle.Format.Line);
					title.Position = chartTitle.Position;
					title.IncludeInLayout = chartTitle.IncludeInLayout;
					title.Top = chartTitle.Top;
					title.Left = chartTitle.Left;
					chartTitle = null;
					title.Font = fontProperties;
					title.Fill = fillProperties;
					title.Border = borderProperties;
				}
				Title = title;
				DataTableProperties dataTable = default(DataTableProperties);
				if (cht.HasDataTable)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
					DataTable dataTable2 = cht.DataTable;
					FontProperties fontProperties = GetFontProperties(dataTable2.Font);
					fillProperties = GetFillProperties(dataTable2.Format.Fill);
					borderProperties = GetBorderProperties(dataTable2.Format.Line);
					dataTable.ShowLegendKey = dataTable2.ShowLegendKey;
					dataTable.HasBorderHorizontal = dataTable2.HasBorderHorizontal;
					dataTable.HasBorderOutline = dataTable2.HasBorderOutline;
					dataTable.HasBorderVertical = dataTable2.HasBorderVertical;
					dataTable2 = null;
					dataTable.Font = fontProperties;
					dataTable.Fill = fillProperties;
					dataTable.Border = borderProperties;
				}
				DataTable = dataTable;
				AxisProperties primaryValueAxis = default(AxisProperties);
				if (Chart.HasPrimaryValueAxis)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
					primaryValueAxis = A((Axis)cht.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue));
				}
				PrimaryValueAxis = primaryValueAxis;
				primaryValueAxis = default(AxisProperties);
				if (Chart.HasPrimaryCategoryAxis)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
					primaryValueAxis = A((Axis)cht.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory));
				}
				PrimaryCategoryAxis = primaryValueAxis;
				primaryValueAxis = default(AxisProperties);
				if (Chart.HasSecondaryValueAxis)
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
					primaryValueAxis = A((Axis)cht.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue, XlAxisGroup.xlSecondary));
				}
				SecondaryValueAxis = primaryValueAxis;
				primaryValueAxis = default(AxisProperties);
				if (Chart.HasSecondaryCategoryAxis)
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
					primaryValueAxis = A((Axis)cht.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, XlAxisGroup.xlSecondary));
				}
				SecondaryCategoryAxis = primaryValueAxis;
				try
				{
					ChartGroups chartGroups2 = (ChartGroups)cht.ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value));
					int count3 = chartGroups2.Count;
					for (int k = 1; k <= count3; k++)
					{
						try
						{
							dictionary6.Add(k, chartGroups2.Item(k).Overlap);
						}
						catch (Exception ex21)
						{
							ProjectData.SetProjectError(ex21);
							Exception ex22 = ex21;
							ProjectData.ClearProjectError();
						}
						try
						{
							if (!A(cht.ChartType))
							{
								continue;
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								dictionary7.Add(k, chartGroups2.Item(k).FirstSliceAngle);
								break;
							}
						}
						catch (Exception ex23)
						{
							ProjectData.SetProjectError(ex23);
							Exception ex24 = ex23;
							ProjectData.ClearProjectError();
						}
					}
					chartGroups2 = null;
				}
				catch (Exception ex25)
				{
					ProjectData.SetProjectError(ex25);
					Exception ex26 = ex25;
					ProjectData.ClearProjectError();
				}
				string filename = modFunctionsIO.PathGetTempFileName();
				try
				{
					cht.Export(filename, VH.A(125592), RuntimeHelpers.GetObjectValue(Missing.Value));
					Bitmap bitmap = new Bitmap(filename);
					BitmapSource imageSource;
					try
					{
						imageSource = Forms.GetImageSource(bitmap);
					}
					finally
					{
						if (bitmap != null)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								((IDisposable)bitmap).Dispose();
								break;
							}
						}
					}
					SourceImage = imageSource;
				}
				catch (Exception ex27)
				{
					ProjectData.SetProjectError(ex27);
					Exception ex28 = ex27;
					Interaction.MsgBox(ex28.Message);
					SourceImage = null;
					ProjectData.ClearProjectError();
				}
			}
			catch (Exception ex29)
			{
				ProjectData.SetProjectError(ex29);
				Exception ex30 = ex29;
				MessageBox.Show(VH.A(172932) + ex30.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
				clsReporting.LogException(ex30);
				ProjectData.ClearProjectError();
			}
		}
	}

	private bool A(XlChartType A)
	{
		if (A <= XlChartType.xl3DPie)
		{
			while (true)
			{
				switch (3)
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
			if (A != XlChartType.xlDoughnut)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				if (A != XlChartType.xl3DPie)
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
					goto IL_0075;
				}
			}
		}
		else if (A != XlChartType.xlPie)
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
			if ((uint)(A - 68) > 3u)
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
				if (A != XlChartType.xlDoughnutExploded)
				{
					goto IL_0075;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
			}
		}
		return true;
		IL_0075:
		return false;
	}

	public static FontProperties GetFontProperties(Font2 font)
	{
		FontProperties result = default(FontProperties);
		Font2 font2 = font;
		result.Size = font2.Size;
		result.ForeColor = font2.Fill.ForeColor.RGB;
		result.BackColor = font2.Fill.ForeColor.RGB;
		result.Name = font2.Name;
		result.Decoration.Bold = font2.Bold;
		result.Decoration.Italic = font2.Italic;
		result.Decoration.UnderlineStyle = font2.UnderlineStyle;
		font2 = null;
		return result;
	}

	public static FontProperties GetFontProperties(Microsoft.Office.Interop.Excel.Font font)
	{
		FontProperties result = default(FontProperties);
		Microsoft.Office.Interop.Excel.Font font2 = font;
		result.Size = Conversions.ToSingle(font2.Size);
		result.ForeColor = Conversions.ToInteger(font2.Color);
		result.BackColor = Conversions.ToInteger(font2.Color);
		result.Name = Conversions.ToString(font2.Name);
		result.Decoration.Bold = (MsoTriState)Conversions.ToInteger(font2.Bold);
		result.Decoration.Italic = (MsoTriState)Conversions.ToInteger(font2.Italic);
		XlUnderlineStyle xlUnderlineStyle = (XlUnderlineStyle)font2.Underline;
		if (xlUnderlineStyle != XlUnderlineStyle.xlUnderlineStyleNone)
		{
			if (xlUnderlineStyle != XlUnderlineStyle.xlUnderlineStyleDouble)
			{
				while (true)
				{
					switch (3)
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
				switch (xlUnderlineStyle)
				{
				case XlUnderlineStyle.xlUnderlineStyleSingle:
				case XlUnderlineStyle.xlUnderlineStyleSingleAccounting:
					result.Decoration.UnderlineStyle = MsoTextUnderlineType.msoUnderlineSingleLine;
					goto IL_010c;
				case XlUnderlineStyle.xlUnderlineStyleDoubleAccounting:
					break;
				default:
					goto IL_010c;
				}
			}
			result.Decoration.UnderlineStyle = MsoTextUnderlineType.msoUnderlineDoubleLine;
		}
		else
		{
			result.Decoration.UnderlineStyle = MsoTextUnderlineType.msoNoUnderline;
		}
		goto IL_010c;
		IL_010c:
		font2 = null;
		return result;
	}

	public static FillProperties GetFillProperties(Microsoft.Office.Interop.Excel.FillFormat fill)
	{
		FillProperties result = default(FillProperties);
		Microsoft.Office.Interop.Excel.FillFormat fillFormat = fill;
		result.BackColor = fillFormat.BackColor.RGB;
		result.ForeColor = fillFormat.ForeColor.RGB;
		result.ObjectThemeColor = fillFormat.ForeColor.ObjectThemeColor;
		result.Pattern = fillFormat.Pattern;
		result.Transparency = Math.Max(0f, fillFormat.Transparency);
		result.Type = fillFormat.Type;
		result.Visible = fillFormat.Visible;
		if (fillFormat.Type == MsoFillType.msoFillGradient)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			result.GradientAngle = fillFormat.GradientAngle;
			if (fillFormat.GradientColorType == MsoGradientColorType.msoGradientOneColor)
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
				result.GradientDegree = fillFormat.GradientDegree;
			}
			List<GradientStopProperties> list = new List<GradientStopProperties>();
			int count = fillFormat.GradientStops.Count;
			for (int i = 1; i <= count; i = checked(i + 1))
			{
				GradientStopProperties item = default(GradientStopProperties);
				GradientStop gradientStop = fillFormat.GradientStops[i];
				Microsoft.Office.Core.ColorFormat color = gradientStop.Color;
				item.RGB = color.RGB;
				item.SchemeColor = color.SchemeColor;
				item.Brightness = color.Brightness;
				item.Type = color.Type;
				color = null;
				item.ObjectThemeColor = gradientStop.Color.ObjectThemeColor;
				item.Transparency = gradientStop.Transparency;
				item.Position = gradientStop.Position;
				gradientStop = null;
				list.Add(item);
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			result.GradientStops = list;
			list = null;
			result.GradientStyle = fillFormat.GradientStyle;
			result.GradientColorType = fillFormat.GradientColorType;
			result.GradientVariant = fillFormat.GradientVariant;
			if (fillFormat.GradientColorType == MsoGradientColorType.msoGradientPresetColors)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					break;
				}
				result.PresetGradientType = fillFormat.PresetGradientType;
			}
		}
		fillFormat = null;
		return result;
	}

	public static LineProperties GetBorderProperties(LineFormat line)
	{
		LineProperties result = default(LineProperties);
		LineFormat lineFormat = line;
		result.ForeColor = lineFormat.ForeColor.RGB;
		result.DashStyle = lineFormat.DashStyle;
		result.Style = lineFormat.Style;
		result.Transparency = lineFormat.Transparency;
		if (lineFormat.Weight < 0f)
		{
			while (true)
			{
				switch (4)
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
			result.Weight = 0f;
			result.Visible = MsoTriState.msoFalse;
		}
		else
		{
			result.Weight = lineFormat.Weight;
			result.Visible = lineFormat.Visible;
		}
		lineFormat = null;
		return result;
	}

	public static BorderProperties GetBorderProperties(Border border)
	{
		BorderProperties result = default(BorderProperties);
		Border border2 = border;
		result.Color = Conversions.ToInteger(border2.Color);
		result.LineStyle = (XlLineStyle)Conversions.ToInteger(border2.LineStyle);
		result.Weight = Conversions.ToSingle(border2.Weight);
		border2 = null;
		return result;
	}

	private AxisProperties A(Axis A)
	{
		AxisProperties result = default(AxisProperties);
		Axis axis = A;
		if (A.Type == Microsoft.Office.Interop.Excel.XlAxisType.xlValue)
		{
			result.MaximumScale = axis.MaximumScale;
			result.MaximumScaleIsAuto = axis.MaximumScaleIsAuto;
			result.MinimumScale = (float)axis.MinimumScale;
			result.MinimumScaleIsAuto = axis.MinimumScaleIsAuto;
			result.MajorUnit = axis.MajorUnit;
			result.MajorUnitIsAuto = axis.MajorUnitIsAuto;
			result.MinorUnit = axis.MajorUnit;
			result.MinorUnitIsAuto = axis.MajorUnitIsAuto;
			if (axis.DisplayUnit != (XlDisplayUnit)(-4142))
			{
				while (true)
				{
					switch (6)
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
				result.DisplayUnit = axis.DisplayUnit;
				result.HasDisplayUnitLabel = axis.HasDisplayUnitLabel;
			}
			result.ScaleType = axis.ScaleType;
		}
		else
		{
			try
			{
				result.AxisBetweenCategories = axis.AxisBetweenCategories;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			try
			{
				result.MajorUnitScale = axis.MajorUnitScale;
				result.MinorUnitScale = axis.MinorUnitScale;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		result.Crosses = axis.Crosses;
		if (axis.Crosses == XlAxisCrosses.xlAxisCrossesCustom)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			result.CrossesAt = axis.CrossesAt;
		}
		result.ReversePlotOrder = axis.ReversePlotOrder;
		result.HasMajorGridlines = axis.HasMajorGridlines;
		result.HasMinorGridlines = axis.HasMinorGridlines;
		if (axis.HasMajorGridlines)
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
			result.MajorGridlines = new GridlinesProperties
			{
				Format = GetBorderProperties(axis.MajorGridlines.Format.Line)
			};
		}
		if (axis.HasMinorGridlines)
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
			result.MinorGridlines = new GridlinesProperties
			{
				Format = GetBorderProperties(axis.MinorGridlines.Format.Line)
			};
		}
		result.Border = GetBorderProperties(A.Border);
		result.MajorTickMark = axis.MajorTickMark;
		result.MinorTickMark = axis.MinorTickMark;
		if (A.Type != Microsoft.Office.Interop.Excel.XlAxisType.xlValue)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			result.TickMarkSpacing = axis.TickMarkSpacing;
		}
		if (axis.TickLabelPosition != XlTickLabelPosition.xlTickLabelPositionNone)
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
			result.TickLabelPosition = axis.TickLabelPosition;
			if (A.Type != Microsoft.Office.Interop.Excel.XlAxisType.xlValue)
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
				result.TickLabelSpacing = axis.TickLabelSpacing;
				result.TickLabelSpacingIsAuto = axis.TickLabelSpacingIsAuto;
			}
			TickLabelsProperties tickLabels = default(TickLabelsProperties);
			TickLabels tickLabels2 = axis.TickLabels;
			try
			{
				tickLabels.Format = GetFontProperties(tickLabels2.Format.TextFrame2.TextRange.Font);
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				tickLabels.Format = GetFontProperties(tickLabels2.Font);
				ProjectData.ClearProjectError();
			}
			try
			{
				tickLabels.MultiLevel = tickLabels2.MultiLevel;
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
			try
			{
				tickLabels.Offset = tickLabels2.Offset;
			}
			catch (Exception ex9)
			{
				ProjectData.SetProjectError(ex9);
				Exception ex10 = ex9;
				ProjectData.ClearProjectError();
			}
			tickLabels.Orientation = (int)tickLabels2.Orientation;
			tickLabels.NumberFormatLinked = tickLabels2.NumberFormatLinked;
			tickLabels.NumberFormat = tickLabels2.NumberFormat;
			try
			{
				tickLabels.Alignment = tickLabels2.Alignment;
			}
			catch (Exception ex11)
			{
				ProjectData.SetProjectError(ex11);
				Exception ex12 = ex11;
				ProjectData.ClearProjectError();
			}
			tickLabels2 = null;
			result.TickLabels = tickLabels;
		}
		result.HasTitle = axis.HasTitle;
		if (axis.HasTitle)
		{
			AxisTitleProperties axisTitle = default(AxisTitleProperties);
			AxisTitle axisTitle2 = axis.AxisTitle;
			try
			{
				axisTitle.Font = GetFontProperties(axisTitle2.Format.TextFrame2.TextRange.Font);
			}
			catch (Exception ex13)
			{
				ProjectData.SetProjectError(ex13);
				Exception ex14 = ex13;
				axisTitle.Font = GetFontProperties(axisTitle2.Font);
				ProjectData.ClearProjectError();
			}
			axisTitle.Fill = GetFillProperties(axisTitle2.Format.Fill);
			axisTitle.Border = GetBorderProperties(axisTitle2.Format.Line);
			axisTitle.IncludeInLayout = axisTitle2.IncludeInLayout;
			axisTitle.HorizontalAlignment = RuntimeHelpers.GetObjectValue(axisTitle2.HorizontalAlignment);
			axisTitle.VerticalAlignment = RuntimeHelpers.GetObjectValue(axisTitle2.VerticalAlignment);
			axisTitle.Orientation = RuntimeHelpers.GetObjectValue(axisTitle2.Orientation);
			axisTitle2 = null;
			result.AxisTitle = axisTitle;
		}
		axis = null;
		return result;
	}
}
