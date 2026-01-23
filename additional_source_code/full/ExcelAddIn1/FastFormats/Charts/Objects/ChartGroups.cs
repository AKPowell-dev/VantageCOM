using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class ChartGroups : ChartItem
{
	[CompilerGenerated]
	private new string m_A;

	[CompilerGenerated]
	private new int? m_A;

	[CompilerGenerated]
	private new int? m_B;

	[CompilerGenerated]
	private new int? C;

	[CompilerGenerated]
	private int? D;

	[CompilerGenerated]
	private int? E;

	[CompilerGenerated]
	private new bool? m_A;

	[CompilerGenerated]
	private new XlSizeRepresents? m_A;

	[CompilerGenerated]
	private int? F;

	[CompilerGenerated]
	private new bool? m_B;

	[CompilerGenerated]
	private new bool? C;

	[CompilerGenerated]
	private new UpDownBars m_A;

	[CompilerGenerated]
	private new UpDownBars m_B;

	[CompilerGenerated]
	private bool? D;

	[CompilerGenerated]
	private new HiLoLines m_A;

	[CompilerGenerated]
	private bool? E;

	[CompilerGenerated]
	private new DropLines m_A;

	[CompilerGenerated]
	private bool? F;

	[CompilerGenerated]
	private new SeriesLines m_A;

	[CompilerGenerated]
	private bool? G;

	[CompilerGenerated]
	private bool? H;

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

	private int? _gapWidth
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

	private int? _overlap
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

	private int? _firstSliceAngle
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

	private int? _donutHoleSize
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

	private int? _bubbleScale
	{
		[CompilerGenerated]
		get
		{
			return this.E;
		}
		[CompilerGenerated]
		set
		{
			this.E = value;
		}
	}

	private bool? _showNegativeBubbles
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

	private XlSizeRepresents? _sizeRepresents
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

	private int? _secondPlotSize
	{
		[CompilerGenerated]
		get
		{
			return this.F;
		}
		[CompilerGenerated]
		set
		{
			this.F = value;
		}
	}

	private bool? _varyByCategories
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

	private bool? _hasUpDownBars
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

	private UpDownBars _upBars
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

	private UpDownBars _downBars
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

	private bool? _hasHiLoLines
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

	private HiLoLines _hiLoLines
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

	private bool? _hasDropLines
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[CompilerGenerated]
		set
		{
			E = value;
		}
	}

	private DropLines _dropLines
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

	private bool? _hasSeriesLines
	{
		[CompilerGenerated]
		get
		{
			return F;
		}
		[CompilerGenerated]
		set
		{
			F = value;
		}
	}

	private SeriesLines _seriesLines
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

	private bool? _has3DShading
	{
		[CompilerGenerated]
		get
		{
			return G;
		}
		[CompilerGenerated]
		set
		{
			G = value;
		}
	}

	private bool? _hasRadarAxisLabels
	{
		[CompilerGenerated]
		get
		{
			return H;
		}
		[CompilerGenerated]
		set
		{
			H = value;
		}
	}

	public bool IsBaseNode => Operators.CompareString(_chartNodeType, "", TextCompare: false) == 0;

	public ChartGroups(XmlNode nd, string chartNodeType)
	{
		_gapWidth = null;
		_overlap = null;
		_firstSliceAngle = null;
		_donutHoleSize = null;
		_bubbleScale = null;
		_showNegativeBubbles = null;
		_sizeRepresents = null;
		_secondPlotSize = null;
		_varyByCategories = null;
		_hasUpDownBars = null;
		_hasHiLoLines = null;
		_hasDropLines = null;
		_hasSeriesLines = null;
		_has3DShading = null;
		_hasRadarAxisLabels = null;
		_chartNodeType = chartNodeType;
		string attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_GAP_WIDTH);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			_gapWidth = Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_OVERLAP);
		if (attributeValue.Length > 0)
		{
			_overlap = Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_FIRST_SLICE_ANGLE);
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
			_firstSliceAngle = Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_DONUT_HOLE_SIZE);
		if (attributeValue.Length > 0)
		{
			_donutHoleSize = Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_BUBBLE_SCALE);
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
			_bubbleScale = Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_SHOW_NEG_BUBBLES);
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
			_showNegativeBubbles = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_SIZE_REPRESENTS);
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
			_sizeRepresents = (XlSizeRepresents)Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_SECOND_PLOT_SIZE);
		if (attributeValue.Length > 0)
		{
			_secondPlotSize = Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_VARY_BY_CAT);
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
			_varyByCategories = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_HAS_UP_DOWN_BARS);
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
			_hasUpDownBars = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_HAS_HI_LO_LINES);
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
			_hasHiLoLines = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_HAS_DROP_LINES);
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
			_hasDropLines = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_HAS_SERIES_LINES);
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
			_hasSeriesLines = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_HAS_3D_SHADING);
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
			_has3DShading = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_HAS_RADAR_AXIS_LABELS);
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
			_hasRadarAxisLabels = Conversions.ToBoolean(attributeValue);
		}
		_upBars = new UpDownBars(nd.SelectSingleNode(FormatConstants.NODE_UP_BARS));
		_downBars = new UpDownBars(nd.SelectSingleNode(FormatConstants.NODE_DOWN_BARS));
		_hiLoLines = new HiLoLines(nd.SelectSingleNode(FormatConstants.NODE_HILO_LINES));
		_dropLines = new DropLines(nd.SelectSingleNode(FormatConstants.NODE_DROP_LINES));
		_seriesLines = new SeriesLines(nd.SelectSingleNode(FormatConstants.NODE_SERIES_LINES));
	}

	private void B(ChartGroup A)
	{
		try
		{
			if (!IsBaseNode)
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
				if (!Operators.ConditionalCompareObjectEqual(Conversions.ToInteger(_chartNodeType), NewLateBinding.LateGet(A.SeriesCollection(1), null, VH.A(141243), new object[0], null, null, null), TextCompare: false))
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
					break;
				}
			}
			ChartGroup chartGroup = A;
			if (!ChartTypes.IsPieChart((XlChartType)Conversions.ToInteger(NewLateBinding.LateGet(A.SeriesCollection(1), null, VH.A(141243), new object[0], null, null, null))))
			{
				if (_gapWidth.HasValue)
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
					try
					{
						chartGroup.GapWidth = _gapWidth.Value;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				if (_overlap.HasValue)
				{
					try
					{
						chartGroup.Overlap = _overlap.Value;
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
				}
				if (_varyByCategories.HasValue)
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
						chartGroup.VaryByCategories = _varyByCategories.Value;
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						ProjectData.ClearProjectError();
					}
				}
			}
			if (_firstSliceAngle.HasValue)
			{
				try
				{
					chartGroup.FirstSliceAngle = _firstSliceAngle.Value;
				}
				catch (Exception ex7)
				{
					ProjectData.SetProjectError(ex7);
					Exception ex8 = ex7;
					ProjectData.ClearProjectError();
				}
			}
			if (_donutHoleSize.HasValue)
			{
				try
				{
					chartGroup.DoughnutHoleSize = _donutHoleSize.Value;
				}
				catch (Exception ex9)
				{
					ProjectData.SetProjectError(ex9);
					Exception ex10 = ex9;
					ProjectData.ClearProjectError();
				}
			}
			if (_bubbleScale.HasValue)
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
					chartGroup.BubbleScale = _bubbleScale.Value;
				}
				catch (Exception ex11)
				{
					ProjectData.SetProjectError(ex11);
					Exception ex12 = ex11;
					ProjectData.ClearProjectError();
				}
			}
			if (_showNegativeBubbles.HasValue)
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
					chartGroup.ShowNegativeBubbles = _showNegativeBubbles.Value;
				}
				catch (Exception ex13)
				{
					ProjectData.SetProjectError(ex13);
					Exception ex14 = ex13;
					ProjectData.ClearProjectError();
				}
			}
			if (_sizeRepresents.HasValue)
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
					chartGroup.SizeRepresents = _sizeRepresents.Value;
				}
				catch (Exception ex15)
				{
					ProjectData.SetProjectError(ex15);
					Exception ex16 = ex15;
					ProjectData.ClearProjectError();
				}
			}
			if (_secondPlotSize.HasValue)
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
					chartGroup.SecondPlotSize = _secondPlotSize.Value;
				}
				catch (Exception ex17)
				{
					ProjectData.SetProjectError(ex17);
					Exception ex18 = ex17;
					ProjectData.ClearProjectError();
				}
			}
			try
			{
				if (_hasUpDownBars.HasValue)
				{
					chartGroup.HasUpDownBars = _hasUpDownBars.Value;
				}
				if (chartGroup.HasUpDownBars)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						_upBars.A(chartGroup.UpBars.Format);
						_downBars.A(chartGroup.DownBars.Format);
						break;
					}
				}
			}
			catch (Exception ex19)
			{
				ProjectData.SetProjectError(ex19);
				Exception ex20 = ex19;
				ProjectData.ClearProjectError();
			}
			try
			{
				if (_hasHiLoLines.HasValue)
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
					chartGroup.HasHiLoLines = _hasHiLoLines.Value;
				}
				if (chartGroup.HasHiLoLines)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						_hiLoLines.A(chartGroup.HiLoLines);
						break;
					}
				}
			}
			catch (Exception ex21)
			{
				ProjectData.SetProjectError(ex21);
				Exception ex22 = ex21;
				ProjectData.ClearProjectError();
			}
			try
			{
				if (_hasDropLines.HasValue)
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
					chartGroup.HasDropLines = _hasDropLines.Value;
				}
				if (chartGroup.HasDropLines)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						_dropLines.A(chartGroup.DropLines);
						break;
					}
				}
			}
			catch (Exception ex23)
			{
				ProjectData.SetProjectError(ex23);
				Exception ex24 = ex23;
				ProjectData.ClearProjectError();
			}
			try
			{
				if (_hasSeriesLines.HasValue)
				{
					chartGroup.HasSeriesLines = _hasSeriesLines.Value;
				}
				if (chartGroup.HasSeriesLines)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						_seriesLines.A(chartGroup.SeriesLines);
						break;
					}
				}
			}
			catch (Exception ex25)
			{
				ProjectData.SetProjectError(ex25);
				Exception ex26 = ex25;
				ProjectData.ClearProjectError();
			}
			if (_has3DShading.HasValue)
			{
				try
				{
					chartGroup.Has3DShading = _has3DShading.Value;
				}
				catch (Exception ex27)
				{
					ProjectData.SetProjectError(ex27);
					Exception ex28 = ex27;
					ProjectData.ClearProjectError();
				}
			}
			try
			{
				if (_hasRadarAxisLabels.HasValue)
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
					chartGroup.HasRadarAxisLabels = _hasRadarAxisLabels.Value;
				}
				_ = chartGroup.HasRadarAxisLabels;
			}
			catch (Exception ex29)
			{
				ProjectData.SetProjectError(ex29);
				Exception ex30 = ex29;
				ProjectData.ClearProjectError();
			}
			chartGroup = null;
		}
		catch (Exception ex31)
		{
			ProjectData.SetProjectError(ex31);
			Exception ex32 = ex31;
			ProjectData.ClearProjectError();
		}
	}

	internal override void A(Microsoft.Office.Interop.Excel.Chart A)
	{
		try
		{
			IEnumerator enumerator = ((IEnumerable)A.ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
					B((ChartGroup)objectValue);
				}
				while (true)
				{
					switch (5)
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
				IDisposable disposable = enumerator as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			object objectValue = null;
			ProjectData.ClearProjectError();
		}
	}
}
