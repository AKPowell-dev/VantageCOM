using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class Chart : ChartItem
{
	[CompilerGenerated]
	private new string m_A;

	[CompilerGenerated]
	private new Dictionary<XlChartType, List<Microsoft.Office.Interop.Excel.Series>> m_A;

	[CompilerGenerated]
	private new double? m_A;

	[CompilerGenerated]
	private new double? m_B;

	[CompilerGenerated]
	private new MsoTriState? m_A;

	[CompilerGenerated]
	private new List<string> m_A;

	[CompilerGenerated]
	private new int m_A;

	[CompilerGenerated]
	private new Font m_A;

	[CompilerGenerated]
	private new ChartArea m_A;

	[CompilerGenerated]
	private new PlotArea m_A;

	[CompilerGenerated]
	private new ChartTitle m_A;

	[CompilerGenerated]
	private new Legend m_A;

	[CompilerGenerated]
	private new DataTable m_A;

	[CompilerGenerated]
	private new Axes m_A;

	[CompilerGenerated]
	private new Axes m_B;

	[CompilerGenerated]
	private new Axes m_C;

	[CompilerGenerated]
	private Axes m_D;

	[CompilerGenerated]
	private Axes m_E;

	[CompilerGenerated]
	private new Series m_A;

	[CompilerGenerated]
	private new ChartGroups m_A;

	[CompilerGenerated]
	private new string m_B;

	[CompilerGenerated]
	private new bool? m_A;

	[CompilerGenerated]
	private new XlCategoryLabelLevel? m_A;

	[CompilerGenerated]
	private new int? m_A;

	[CompilerGenerated]
	private new XlDisplayBlanksAs? m_A;

	[CompilerGenerated]
	private new int? m_B;

	[CompilerGenerated]
	private new int? m_C;

	[CompilerGenerated]
	private int? m_D;

	[CompilerGenerated]
	private new bool? m_B;

	[CompilerGenerated]
	private new bool? m_C;

	[CompilerGenerated]
	private new double? m_C;

	[CompilerGenerated]
	private new XlSeriesNameLevel? m_A;

	[CompilerGenerated]
	private bool? m_D;

	private string _chartNodeType
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private Dictionary<XlChartType, List<Microsoft.Office.Interop.Excel.Series>> _seriesGroupDictionary
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private double? _height
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private double? _width
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	private MsoTriState? _lockAspectRatio
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private List<string> _seriesColors
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private int _seriesColorsCounter
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private Font _baseFont
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private ChartArea _chartArea
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private PlotArea _plotArea
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private ChartTitle _chartTitle
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private Legend _legend
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private DataTable _dataTable
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private Axes _axesCommon
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private Axes _axesPrimaryCategory
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	private Axes _axesPrimaryValue
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	private Axes _axesSecondaryCategory
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	private Axes _axesSecondaryValue
	{
		[CompilerGenerated]
		get
		{
			return this.m_E;
		}
		[CompilerGenerated]
		set
		{
			this.m_E = value;
		}
	}

	private Series _series
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private ChartGroups _chartGroups
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private string _units
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	private bool? _autoScaling
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private XlCategoryLabelLevel? _categoryLabelLevel
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private int? _depthPercent
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private XlDisplayBlanksAs? _displayBlanksAs
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private int? _elevation
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	private int? _heightPercent
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	private int? _perspective
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	private bool? _plotVisibleOnly
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	private bool? _rightAngleAxes
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	private double? _rotation
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	private XlSeriesNameLevel? _seriesNameLevel
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private bool? _showDataLabelsOverMax
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	public bool IsBaseNode => Operators.CompareString(_chartNodeType, "", TextCompare: false) == 0;

	public Chart(XmlNode node)
	{
		_seriesGroupDictionary = null;
		_height = null;
		_width = null;
		_lockAspectRatio = null;
		_seriesColors = null;
		_autoScaling = null;
		_categoryLabelLevel = null;
		_depthPercent = null;
		_displayBlanksAs = null;
		_elevation = null;
		_heightPercent = null;
		_perspective = null;
		_plotVisibleOnly = null;
		_rightAngleAxes = null;
		_rotation = null;
		_seriesNameLevel = null;
		_showDataLabelsOverMax = null;
		XmlNode xmlNode = node.SelectSingleNode(FormatConstants.NODE_BASICS);
		XmlNode xmlNode2 = node.SelectSingleNode(FormatConstants.NODE_BASICS + VH.A(75498) + FormatConstants.NODE_SIZE);
		XmlNodeList xmlNodeList = node.SelectNodes(FormatConstants.NODE_BASICS + VH.A(75498) + FormatConstants.NODE_SERIES_COLORS + VH.A(75498) + FormatConstants.NODE_COLOR);
		XmlNode xmlNode3 = node.SelectSingleNode(FormatConstants.NODE_CHART);
		_chartNodeType = node.Attributes[FormatConstants.ATTR_TYPE].Value;
		_units = xmlNode.Attributes[FormatConstants.ATTR_UNITS].Value;
		if (xmlNodeList != null)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (xmlNodeList.Count > 0)
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
				_seriesColors = new List<string>();
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = xmlNodeList.GetEnumerator();
					while (enumerator.MoveNext())
					{
						string value = ((XmlNode)enumerator.Current).Attributes[FormatConstants.ATTR_FILL_FORE_COLOR].Value;
						_seriesColors.Add(value);
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_0263;
						}
						continue;
						end_IL_0263:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
			}
		}
		xmlNodeList = null;
		string attributeValue = clsXml.GetAttributeValue(xmlNode2, FormatConstants.ATTR_HEIGHT);
		if (attributeValue.Length > 0)
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
			_height = modFunctionsConvert.CvtInvariantStrToDbl(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(xmlNode2, FormatConstants.ATTR_WIDTH);
		if (attributeValue.Length > 0)
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
			_width = modFunctionsConvert.CvtInvariantStrToDbl(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(xmlNode2, FormatConstants.ATTR_LOCK_ASPECT_RATIO);
		if (attributeValue.Length > 0)
		{
			if (Conversions.ToBoolean(attributeValue))
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
				_lockAspectRatio = MsoTriState.msoTrue;
			}
			else
			{
				_lockAspectRatio = MsoTriState.msoFalse;
			}
		}
		XmlNode xmlNode4 = node;
		_baseFont = new Font(xmlNode4.SelectSingleNode(FormatConstants.NODE_BASICS + VH.A(75498) + FormatConstants.NODE_BASE_FONT));
		_chartArea = new ChartArea(xmlNode4.SelectSingleNode(FormatConstants.NODE_CHART_AREA));
		_plotArea = new PlotArea(xmlNode4.SelectSingleNode(FormatConstants.NODE_PLOT_AREA), _units);
		_chartTitle = new ChartTitle(xmlNode4.SelectSingleNode(FormatConstants.NODE_CHART_TITLE), _units);
		_legend = new Legend(xmlNode4.SelectSingleNode(FormatConstants.NODE_LEGEND), _units);
		_dataTable = new DataTable(xmlNode4.SelectSingleNode(FormatConstants.NODE_DATA_TABLE));
		_axesCommon = new Axes(xmlNode4.SelectSingleNode(FormatConstants.NODE_AXES_COMMON), _units);
		_axesPrimaryCategory = new Axes(xmlNode4.SelectSingleNode(FormatConstants.NODE_AXES_PRIMARY_CATEGORY), _units);
		_axesPrimaryValue = new Axes(xmlNode4.SelectSingleNode(FormatConstants.NODE_AXES_PRIMARY_VALUE), _units);
		_axesSecondaryCategory = new Axes(xmlNode4.SelectSingleNode(FormatConstants.NODE_AXES_SECONDARY_CATEGORY), _units);
		_axesSecondaryValue = new Axes(xmlNode4.SelectSingleNode(FormatConstants.NODE_AXES_SECONDARY_VALUE), _units);
		_series = new Series(xmlNode4.SelectSingleNode(FormatConstants.NODE_SERIES), _chartNodeType, _seriesColors);
		_chartGroups = new ChartGroups(xmlNode4.SelectSingleNode(FormatConstants.NODE_CHART_GROUPS), _chartNodeType);
		xmlNode4 = null;
		attributeValue = clsXml.GetAttributeValue(xmlNode3, FormatConstants.ATTR_AUTO_SCALING);
		if (attributeValue.Length > 0)
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
			_autoScaling = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(xmlNode3, FormatConstants.ATTR_CAT_LABEL_LEVEL);
		if (attributeValue.Length > 0)
		{
			_categoryLabelLevel = (XlCategoryLabelLevel)Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(xmlNode3, FormatConstants.ATTR_DEPTH_PCT);
		if (attributeValue.Length > 0)
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
			_depthPercent = Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(xmlNode3, FormatConstants.ATTR_DISP_BLANKS_AS);
		if (attributeValue.Length > 0)
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
			_displayBlanksAs = (XlDisplayBlanksAs)Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(xmlNode3, FormatConstants.ATTR_ELEVATION);
		if (attributeValue.Length > 0)
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
			_elevation = Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(xmlNode3, FormatConstants.ATTR_HEIGHT_PCT);
		if (attributeValue.Length > 0)
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
			_heightPercent = Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(xmlNode3, FormatConstants.ATTR_PERSPECTIVE);
		if (attributeValue.Length > 0)
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
			_perspective = Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(xmlNode3, FormatConstants.ATTR_PLOT_VIS_ONLY);
		if (attributeValue.Length > 0)
		{
			_plotVisibleOnly = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(xmlNode3, FormatConstants.ATTR_RIGHT_ANGLE_AXES);
		if (attributeValue.Length > 0)
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
			_rightAngleAxes = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(xmlNode3, FormatConstants.ATTR_ROTATION);
		if (attributeValue.Length > 0)
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
			_rotation = modFunctionsConvert.CvtInvariantStrToDbl(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(xmlNode3, FormatConstants.ATTR_SERIES_NAME_LEVEL);
		if (attributeValue.Length > 0)
		{
			_seriesNameLevel = (XlSeriesNameLevel)Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(xmlNode3, FormatConstants.ATTR_SHOW_DL_OVER_MAX);
		if (attributeValue.Length > 0)
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
			_showDataLabelsOverMax = Conversions.ToBoolean(attributeValue);
		}
		xmlNode2 = null;
		xmlNode3 = null;
	}

	private void B(Microsoft.Office.Interop.Excel.Series A, Microsoft.Office.Interop.Excel.Chart B)
	{
		try
		{
			_series.UpdateSeriesBorder(A.Format, B);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void B(Microsoft.Office.Interop.Excel.Series A)
	{
		checked
		{
			try
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = ((IEnumerable)A.Points(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
					while (enumerator.MoveNext())
					{
						Point point = (Point)enumerator.Current;
						B(point.Format);
						_seriesColorsCounter++;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
	}

	private void B(List<Microsoft.Office.Interop.Excel.Series> A, Microsoft.Office.Interop.Excel.Chart B)
	{
		try
		{
			using List<Microsoft.Office.Interop.Excel.Series>.Enumerator enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Excel.Series current = enumerator.Current;
				XlChartType chartCombinedType = ChartTypes.GetChartCombinedType(current.ChartType);
				if (!IsBaseNode)
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
					if (Conversions.ToInteger(_chartNodeType) != (int)ChartTypes.GetChartCombinedType(chartCombinedType))
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
						break;
					}
				}
				try
				{
					if (!ChartTypes.HasLineSeries(chartCombinedType))
					{
						_series.UpdateSeriesBorder(current.Format, B);
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
	}

	private void C(Microsoft.Office.Interop.Excel.Series A)
	{
		A.Format.Line.Visible = MsoTriState.msoTrue;
		C(A.Format);
		A.Format.Line.Visible = MsoTriState.msoFalse;
		_series.UpdateSeriesMarkersColor(A, _seriesColors[B()]);
	}

	private void B(List<Microsoft.Office.Interop.Excel.Series> A)
	{
		try
		{
			foreach (Microsoft.Office.Interop.Excel.Series item in A)
			{
				XlChartType chartCombinedType = ChartTypes.GetChartCombinedType(item.ChartType);
				if (!IsBaseNode)
				{
					if (Conversions.ToInteger(_chartNodeType) != (int)chartCombinedType)
					{
						continue;
					}
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
				}
				checked
				{
					if (ChartTypes.HasLineSeries(chartCombinedType))
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
						try
						{
							if (item.Format.Line.Visible == MsoTriState.msoTrue)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									C(item.Format);
									_seriesColorsCounter++;
									break;
								}
							}
							else
							{
								if (!Conversions.ToBoolean(FormatUtil.SeriesHasMarkers(item)))
								{
									continue;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									C(item);
									_seriesColorsCounter++;
									break;
								}
								continue;
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						continue;
					}
					try
					{
						if (item.Format.Fill.Visible != MsoTriState.msoTrue)
						{
							continue;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							B(item.Format);
							_seriesColorsCounter++;
							break;
						}
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
				}
			}
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
	}

	private void C(List<Microsoft.Office.Interop.Excel.Series> A)
	{
		if (_seriesColors == null)
		{
			return;
		}
		while (true)
		{
			switch (3)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (!ChartTypes.IsSurfaceChart(A.First().ChartType))
			{
				if (ChartTypes.IsPieChart(A.First().ChartType))
				{
					B(A.First());
				}
				else
				{
					B(A);
				}
			}
			return;
		}
	}

	private void C(List<Microsoft.Office.Interop.Excel.Series> A, Microsoft.Office.Interop.Excel.Chart B)
	{
		if (ChartTypes.IsSurfaceChart(A.First().ChartType))
		{
			return;
		}
		while (true)
		{
			switch (3)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (ChartTypes.IsPieChart(A.First().ChartType))
			{
				this.B(A.First(), B);
			}
			else
			{
				this.B(A, B);
			}
			return;
		}
	}

	private void D(List<Microsoft.Office.Interop.Excel.Series> A)
	{
		try
		{
			using List<Microsoft.Office.Interop.Excel.Series>.Enumerator enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Excel.Series current = enumerator.Current;
				XlChartType chartCombinedType = ChartTypes.GetChartCombinedType(current.ChartType);
				if ((!IsBaseNode && Conversions.ToInteger(_chartNodeType) != (int)chartCombinedType) || !ChartTypes.HasLineSeries(chartCombinedType))
				{
					continue;
				}
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
				try
				{
					if (current.MarkerStyle == XlMarkerStyle.xlMarkerStyleNone)
					{
						continue;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						_series.UpdateSeriesMarkers(current);
						break;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
	}

	private void B(Microsoft.Office.Interop.Excel.Chart A)
	{
		ChartObject chartObject = (ChartObject)A.Parent;
		MsoTriState lockAspectRatio = default(MsoTriState);
		try
		{
			lockAspectRatio = chartObject.ShapeRange.LockAspectRatio;
			chartObject.ShapeRange.LockAspectRatio = MsoTriState.msoFalse;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (_height.HasValue)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			try
			{
				chartObject.Height = FormatUtil.GetDimensionInPoints(_height.Value, _units);
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		if (_width.HasValue)
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
			try
			{
				chartObject.Width = FormatUtil.GetDimensionInPoints(_width.Value, _units);
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
		}
		try
		{
			if (_lockAspectRatio.HasValue)
			{
				chartObject.ShapeRange.LockAspectRatio = _lockAspectRatio.Value;
			}
			else
			{
				chartObject.ShapeRange.LockAspectRatio = lockAspectRatio;
			}
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
		chartObject = null;
	}

	private void C(Microsoft.Office.Interop.Excel.Chart A)
	{
		Microsoft.Office.Interop.Excel.Chart chart = A;
		if (_autoScaling.HasValue)
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
			try
			{
				chart.AutoScaling = _autoScaling.Value;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		if (_categoryLabelLevel.HasValue)
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
			try
			{
				chart.CategoryLabelLevel = _categoryLabelLevel.Value;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		if (_depthPercent.HasValue)
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
			try
			{
				chart.DepthPercent = _depthPercent.Value;
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
		}
		if (_displayBlanksAs.HasValue)
		{
			try
			{
				chart.DisplayBlanksAs = _displayBlanksAs.Value;
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
		}
		if (_elevation.HasValue)
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
			try
			{
				chart.Elevation = _elevation.Value;
			}
			catch (Exception ex9)
			{
				ProjectData.SetProjectError(ex9);
				Exception ex10 = ex9;
				ProjectData.ClearProjectError();
			}
		}
		if (_heightPercent.HasValue)
		{
			try
			{
				chart.HeightPercent = _heightPercent.Value;
			}
			catch (Exception ex11)
			{
				ProjectData.SetProjectError(ex11);
				Exception ex12 = ex11;
				ProjectData.ClearProjectError();
			}
		}
		if (_perspective.HasValue)
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
			try
			{
				chart.Perspective = checked(_perspective.Value * 2);
			}
			catch (Exception ex13)
			{
				ProjectData.SetProjectError(ex13);
				Exception ex14 = ex13;
				ProjectData.ClearProjectError();
			}
		}
		if (_plotVisibleOnly.HasValue)
		{
			try
			{
				chart.PlotVisibleOnly = _plotVisibleOnly.Value;
			}
			catch (Exception ex15)
			{
				ProjectData.SetProjectError(ex15);
				Exception ex16 = ex15;
				ProjectData.ClearProjectError();
			}
		}
		if (_rightAngleAxes.HasValue)
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
			try
			{
				chart.RightAngleAxes = _rightAngleAxes.Value;
			}
			catch (Exception ex17)
			{
				ProjectData.SetProjectError(ex17);
				Exception ex18 = ex17;
				ProjectData.ClearProjectError();
			}
		}
		if (_rotation.HasValue)
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
			try
			{
				chart.Rotation = _rotation.Value;
			}
			catch (Exception ex19)
			{
				ProjectData.SetProjectError(ex19);
				Exception ex20 = ex19;
				ProjectData.ClearProjectError();
			}
		}
		if (_seriesNameLevel.HasValue)
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
			try
			{
				chart.SeriesNameLevel = _seriesNameLevel.Value;
			}
			catch (Exception ex21)
			{
				ProjectData.SetProjectError(ex21);
				Exception ex22 = ex21;
				ProjectData.ClearProjectError();
			}
		}
		if (_showDataLabelsOverMax.HasValue)
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
			try
			{
				chart.ShowDataLabelsOverMaximum = _showDataLabelsOverMax.Value;
			}
			catch (Exception ex23)
			{
				ProjectData.SetProjectError(ex23);
				Exception ex24 = ex23;
				ProjectData.ClearProjectError();
			}
		}
		chart = null;
	}

	private void D(Microsoft.Office.Interop.Excel.Chart A)
	{
		B(A);
		C(A);
		_baseFont.A(A);
		_chartGroups.A(A);
		_series.A(A);
		_chartArea.A(A);
		_chartTitle.A(A);
		_legend.A(A);
		_dataTable.A(A);
		_axesPrimaryCategory.A(A);
		_axesPrimaryValue.A(A);
		_axesSecondaryCategory.A(A);
		_axesSecondaryValue.A(A);
		_axesCommon.A(A);
		_plotArea.A(A);
	}

	private void E(Microsoft.Office.Interop.Excel.Chart A)
	{
		_series.A(A);
		_chartGroups.A(A);
	}

	public void ApplyTo(Microsoft.Office.Interop.Excel.Chart selectedChart, Dictionary<XlChartType, List<Microsoft.Office.Interop.Excel.Series>> seriesGroupDictionary)
	{
		_seriesColorsCounter = 0;
		_seriesGroupDictionary = seriesGroupDictionary;
		A(selectedChart);
	}

	internal override void A(Microsoft.Office.Interop.Excel.Chart A)
	{
		if (IsBaseNode)
		{
			D(A);
		}
		else if (A.ChartType != (XlChartType)FormatConstants.CUST_COMBO_CHART_XLCHARTTYPE)
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
			D(A);
		}
		else
		{
			E(A);
		}
		F(A);
	}

	private void F(Microsoft.Office.Interop.Excel.Chart A)
	{
		using Dictionary<XlChartType, List<Microsoft.Office.Interop.Excel.Series>>.Enumerator enumerator = _seriesGroupDictionary.GetEnumerator();
		while (enumerator.MoveNext())
		{
			List<Microsoft.Office.Interop.Excel.Series> value = enumerator.Current.Value;
			C(value);
			C(value, A);
			D(value);
		}
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			return;
		}
	}

	private void B(ChartFormat A)
	{
		B(A.Fill.ForeColor);
	}

	private void C(ChartFormat A)
	{
		B(A.Line.ForeColor);
	}

	private void B(Microsoft.Office.Interop.Excel.ColorFormat A)
	{
		string text = _seriesColors[B()];
		try
		{
			A.RGB = clsColors.RGB2Ole(text);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private int B()
	{
		if (_seriesColorsCounter >= _seriesColors.Count)
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
			_seriesColorsCounter = 0;
		}
		return _seriesColorsCounter;
	}
}
