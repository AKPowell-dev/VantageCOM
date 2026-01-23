using System;
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

public sealed class Series : ChartItem
{
	[CompilerGenerated]
	private new List<string> m_A;

	[CompilerGenerated]
	private new string m_A;

	[CompilerGenerated]
	private new string m_B;

	[CompilerGenerated]
	private new bool? m_A;

	[CompilerGenerated]
	private new DataLabels m_A;

	[CompilerGenerated]
	private new bool? m_B;

	[CompilerGenerated]
	private new LeaderLines m_A;

	[CompilerGenerated]
	private new bool? C;

	[CompilerGenerated]
	private new ErrorBars m_A;

	[CompilerGenerated]
	private new XlMarkerStyle? m_A;

	[CompilerGenerated]
	private new int? m_A;

	[CompilerGenerated]
	private new string C;

	[CompilerGenerated]
	private string D;

	[CompilerGenerated]
	private new int? m_B;

	[CompilerGenerated]
	private new int? C;

	[CompilerGenerated]
	private int? D;

	[CompilerGenerated]
	private bool? D;

	[CompilerGenerated]
	private new XlBarShape? m_A;

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

	private string _seriesBorder
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

	private bool? _hasDataLabels
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

	private DataLabels _dataLabels
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

	private bool? _hasLeaderLines
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

	private LeaderLines _leaderLines
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

	private bool? _hasErrorBars
	{
		[CompilerGenerated]
		get
		{
			return this.C;
		}
		[CompilerGenerated]
		set
		{
			this.C = value;
		}
	}

	private ErrorBars _errorBars
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

	private XlMarkerStyle? _markerStyle
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

	private int? _markerSize
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

	private string _markerForecolorType
	{
		[CompilerGenerated]
		get
		{
			return this.C;
		}
		[CompilerGenerated]
		set
		{
			this.C = value;
		}
	}

	private string _markerBackcolorType
	{
		[CompilerGenerated]
		get
		{
			return this.D;
		}
		[CompilerGenerated]
		set
		{
			this.D = value;
		}
	}

	private int? _markerForecolor
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

	private int? _markerBackcolor
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

	private int? _explosion
	{
		[CompilerGenerated]
		get
		{
			return this.D;
		}
		[CompilerGenerated]
		set
		{
			this.D = value;
		}
	}

	private bool? _smooth
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

	private XlBarShape? _barShape
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

	public bool IsBaseNode => Operators.CompareString(_chartNodeType, "", TextCompare: false) == 0;

	public Series(XmlNode node, string chartNodeType, List<string> seriesColors)
	{
		_seriesColors = null;
		_chartNodeType = null;
		_seriesBorder = null;
		_hasDataLabels = null;
		_hasLeaderLines = null;
		_hasErrorBars = null;
		_markerStyle = null;
		_markerSize = null;
		_markerForecolorType = null;
		_markerBackcolorType = null;
		_markerForecolor = null;
		_markerBackcolor = null;
		_explosion = null;
		_smooth = null;
		_barShape = null;
		_chartNodeType = chartNodeType;
		_seriesColors = seriesColors;
		string attributeValue = clsXml.GetAttributeValue(node, FormatConstants.ATTR_EXPLOSION);
		if (attributeValue.Length > 0)
		{
			_explosion = Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(node, FormatConstants.ATTR_SMOOTH);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			_smooth = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(node, FormatConstants.ATTR_BAR_SHAPE);
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
			_barShape = (XlBarShape)Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(node, FormatConstants.ATTR_HAS_DATA_LABELS);
		if (attributeValue.Length > 0)
		{
			_hasDataLabels = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(node, FormatConstants.ATTR_HAS_LEADER_LINES);
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
			_hasLeaderLines = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(node, FormatConstants.ATTR_HAS_ERROR_BARS);
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
			_hasErrorBars = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(node, FormatConstants.ATTR_SERIES_BORDER);
		if (attributeValue.Length > 0)
		{
			_seriesBorder = attributeValue;
		}
		attributeValue = clsXml.GetAttributeValue(node, FormatConstants.ATTR_MARKER_STYLE);
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
			_markerStyle = (XlMarkerStyle)Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(node, FormatConstants.ATTR_MARKER_SIZE);
		if (attributeValue.Length > 0)
		{
			_markerSize = Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(node, FormatConstants.ATTR_MARKER_FORECOLOR_TYPE);
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
			_markerForecolorType = attributeValue;
		}
		attributeValue = clsXml.GetAttributeValue(node, FormatConstants.ATTR_MARKER_BACKCOLOR_TYPE);
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
			_markerBackcolorType = attributeValue;
		}
		attributeValue = clsXml.GetAttributeValue(node, FormatConstants.ATTR_MARKER_FORECOLOR);
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
			if (Operators.CompareString(attributeValue, FormatConstants.TRANSPARENCY.ToString(), TextCompare: false) != 0)
			{
				_markerForecolor = clsColors.RGB2Ole(attributeValue);
			}
			else
			{
				_markerForecolor = FormatConstants.TRANSPARENCY;
			}
		}
		attributeValue = clsXml.GetAttributeValue(node, FormatConstants.ATTR_MARKER_BACKCOLOR);
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
			if (Operators.CompareString(attributeValue, FormatConstants.TRANSPARENCY.ToString(), TextCompare: false) != 0)
			{
				_markerBackcolor = clsColors.RGB2Ole(attributeValue);
			}
			else
			{
				_markerBackcolor = FormatConstants.TRANSPARENCY;
			}
		}
		XmlNode xmlNode = node;
		_dataLabels = new DataLabels(xmlNode.SelectSingleNode(FormatConstants.NODE_DATA_LABELS));
		_leaderLines = new LeaderLines(xmlNode.SelectSingleNode(FormatConstants.NODE_LEADER_LINES));
		_errorBars = new ErrorBars(xmlNode.SelectSingleNode(FormatConstants.NODE_ERROR_BARS));
		xmlNode = null;
	}

	private void B(FullSeriesCollection A)
	{
		int num = checked(A.Count - 1);
		for (int i = 0; i <= num; i = checked(i + 1))
		{
			Microsoft.Office.Interop.Excel.Series series = (Microsoft.Office.Interop.Excel.Series)A.Cast<object>().ElementAtOrDefault(i);
			try
			{
				if (IsBaseNode)
				{
					goto IL_0066;
				}
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
				if (Conversions.ToInteger(_chartNodeType) == (int)ChartTypes.GetChartCombinedType(series.ChartType))
				{
					goto IL_0066;
				}
				goto end_IL_0028;
				IL_0066:
				if (_explosion.HasValue)
				{
					try
					{
						series.Explosion = _explosion.Value;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				if (_smooth.HasValue)
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
						series.Smooth = _smooth.Value;
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
				}
				if (_barShape.HasValue)
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
						series.BarShape = _barShape.Value;
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
					bool? flag;
					if (_hasDataLabels.HasValue)
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
						bool? hasDataLabels = _hasDataLabels;
						bool? obj;
						if (!hasDataLabels.HasValue)
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
							obj = null;
						}
						else
						{
							obj = hasDataLabels == true;
						}
						flag = obj;
						if (flag.HasValue)
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
							if (flag != true)
							{
								goto IL_01e8;
							}
						}
						if (series.HasDataLabels || !flag.HasValue)
						{
							goto IL_01e8;
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
						series.HasDataLabels = _hasDataLabels.Value;
					}
					goto IL_026e;
					IL_026e:
					if (series.HasDataLabels)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							_dataLabels.A(series);
							break;
						}
					}
					goto end_IL_0130;
					IL_01e8:
					flag = _hasDataLabels;
					bool? obj2;
					if (!flag.HasValue)
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
						obj2 = null;
					}
					else
					{
						obj2 = flag != true;
					}
					flag = obj2;
					if (flag == true)
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
						NewLateBinding.LateCall(series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
					}
					goto IL_026e;
					end_IL_0130:;
				}
				catch (Exception ex7)
				{
					ProjectData.SetProjectError(ex7);
					Exception ex8 = ex7;
					ProjectData.ClearProjectError();
				}
				try
				{
					if (_hasLeaderLines.HasValue)
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
						series.HasLeaderLines = _hasLeaderLines.Value;
					}
					if (series.HasLeaderLines)
					{
						_leaderLines.A(series);
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
					if (_hasErrorBars.HasValue)
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
						series.HasErrorBars = _hasErrorBars.Value;
					}
					if (!series.HasErrorBars)
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
						_errorBars.A(series);
						break;
					}
				}
				catch (Exception ex11)
				{
					ProjectData.SetProjectError(ex11);
					Exception ex12 = ex11;
					ProjectData.ClearProjectError();
				}
				end_IL_0028:;
			}
			catch (Exception ex13)
			{
				ProjectData.SetProjectError(ex13);
				Exception ex14 = ex13;
				ProjectData.ClearProjectError();
			}
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	public void UpdateSeriesBorder(ChartFormat format, Microsoft.Office.Interop.Excel.Chart selectedChart)
	{
		if (string.IsNullOrEmpty(_seriesBorder))
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
			try
			{
				if (string.Equals(_seriesBorder, FormatConstants.NONE))
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							format.Line.Visible = MsoTriState.msoFalse;
							return;
						}
					}
				}
				if (!string.Equals(_seriesBorder, FormatConstants.MATCH_CHART))
				{
					return;
				}
				format.Line.Visible = MsoTriState.msoTrue;
				int chartBackgroundColor = FormatUtil.GetChartBackgroundColor(selectedChart);
				if (chartBackgroundColor <= -1)
				{
					return;
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					format.Line.ForeColor.RGB = chartBackgroundColor;
					return;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
				return;
			}
		}
	}

	public void UpdateSeriesMarkersColor(Microsoft.Office.Interop.Excel.Series series, string color)
	{
		try
		{
			if (series.MarkerStyle == XlMarkerStyle.xlMarkerStyleNone)
			{
				return;
			}
			if (Operators.CompareString(_markerForecolorType, FormatConstants.CUSTOM, TextCompare: false) != 0)
			{
				try
				{
					series.MarkerForegroundColor = clsColors.RGB2Ole(color);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			if (Operators.CompareString(_markerBackcolorType, FormatConstants.CUSTOM, TextCompare: false) == 0)
			{
				return;
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
				try
				{
					series.MarkerBackgroundColor = clsColors.RGB2Ole(color);
					return;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
					return;
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

	public void UpdateSeriesMarkers(Microsoft.Office.Interop.Excel.Series series)
	{
		try
		{
			if (series.MarkerStyle == XlMarkerStyle.xlMarkerStyleNone)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return;
					}
				}
			}
			if (_markerStyle.HasValue)
			{
				try
				{
					series.MarkerStyle = _markerStyle.Value;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			if (_markerSize.HasValue)
			{
				try
				{
					series.MarkerSize = _markerSize.Value;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
			}
			if (Operators.CompareString(_markerForecolorType, FormatConstants.MATCH_SERIES, TextCompare: false) == 0)
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
					if (ChartTypes.HasLineSeries(series.ChartType))
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							if (series.Format.Line.Visible == MsoTriState.msoFalse)
							{
								break;
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								series.MarkerForegroundColor = series.Format.Line.ForeColor.RGB;
								break;
							}
							break;
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
			if (Operators.CompareString(_markerForecolorType, FormatConstants.CUSTOM, TextCompare: false) == 0)
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
				if (_markerForecolor.HasValue)
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
						if (_markerForecolor.Value != FormatConstants.TRANSPARENCY)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								series.MarkerForegroundColor = _markerForecolor.Value;
								break;
							}
						}
						else
						{
							series.MarkerForegroundColorIndex = XlColorIndex.xlColorIndexNone;
						}
					}
					catch (Exception ex7)
					{
						ProjectData.SetProjectError(ex7);
						Exception ex8 = ex7;
						ProjectData.ClearProjectError();
					}
				}
			}
			if (Operators.CompareString(_markerBackcolorType, FormatConstants.MATCH_SERIES, TextCompare: false) == 0)
			{
				try
				{
					if (ChartTypes.HasLineSeries(series.ChartType))
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							if (series.Format.Line.Visible == MsoTriState.msoFalse)
							{
								break;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								series.MarkerBackgroundColor = series.Format.Line.ForeColor.RGB;
								break;
							}
							break;
						}
					}
				}
				catch (Exception ex9)
				{
					ProjectData.SetProjectError(ex9);
					Exception ex10 = ex9;
					ProjectData.ClearProjectError();
				}
			}
			if (Operators.CompareString(_markerBackcolorType, FormatConstants.CUSTOM, TextCompare: false) != 0 || !_markerBackcolor.HasValue)
			{
				return;
			}
			try
			{
				if (_markerBackcolor.Value != FormatConstants.TRANSPARENCY)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							series.MarkerBackgroundColor = _markerBackcolor.Value;
							return;
						}
					}
				}
				series.MarkerBackgroundColorIndex = XlColorIndex.xlColorIndexNone;
			}
			catch (Exception ex11)
			{
				ProjectData.SetProjectError(ex11);
				Exception ex12 = ex11;
				ProjectData.ClearProjectError();
			}
		}
		catch (Exception ex13)
		{
			ProjectData.SetProjectError(ex13);
			Exception ex14 = ex13;
			ProjectData.ClearProjectError();
		}
	}

	private string B(int A)
	{
		if (A < _seriesColors.Count)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return _seriesColors[A];
				}
			}
		}
		return _seriesColors[checked(A - _seriesColors.Count)];
	}

	internal override void A(Microsoft.Office.Interop.Excel.Chart A)
	{
		try
		{
			B((FullSeriesCollection)A.FullSeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
