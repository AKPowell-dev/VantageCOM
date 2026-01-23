using System;
using System.Runtime.CompilerServices;
using System.Xml;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class Axes : ChartItem
{
	[CompilerGenerated]
	private new string m_A;

	[CompilerGenerated]
	private new bool? m_A;

	[CompilerGenerated]
	private new bool? m_B;

	[CompilerGenerated]
	private new Line m_A;

	[CompilerGenerated]
	private new Line m_B;

	[CompilerGenerated]
	private new Line m_C;

	[CompilerGenerated]
	private new AxisTitle m_A;

	[CompilerGenerated]
	private new DisplayUnitLabel m_A;

	[CompilerGenerated]
	private new TickMarks m_A;

	[CompilerGenerated]
	private new TickLabels m_A;

	[CompilerGenerated]
	private new XlDisplayUnit? m_A;

	[CompilerGenerated]
	private new double? m_A;

	[CompilerGenerated]
	private new int? m_A;

	[CompilerGenerated]
	private new bool? m_C;

	[CompilerGenerated]
	private new string m_B;

	private string _nodeName
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

	private bool? _hasMajorGridlines
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

	private bool? _hasMinorGridlines
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

	private Line _majorGridlinesLine
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

	private Line _minorGridlinesLine
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

	private Line _axisLine
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

	private AxisTitle _title
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

	private DisplayUnitLabel _displayUnitLabel
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

	private TickMarks _tickMarks
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

	private TickLabels _tickLabels
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

	private XlDisplayUnit? _displayUnit
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

	private double? _displayUnitCustom
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

	private int? _otherAxisCrossesAt
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

	private bool? _axisBetweenCategories
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

	public Axes(XmlNode nd, string units)
	{
		_nodeName = null;
		_hasMajorGridlines = null;
		_hasMinorGridlines = null;
		_displayUnit = null;
		_displayUnitCustom = null;
		_otherAxisCrossesAt = null;
		_axisBetweenCategories = null;
		_nodeName = nd.Name;
		XmlNode xmlNode = nd.SelectSingleNode(FormatConstants.NODE_MAJOR_GRIDLINES);
		XmlNode xmlNode2 = nd.SelectSingleNode(FormatConstants.NODE_MINOR_GRIDLINES);
		string text = ((BaseItem)this).B(xmlNode);
		if (text.Length > 0)
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
			_hasMajorGridlines = Conversions.ToBoolean(text);
		}
		text = ((BaseItem)this).B(xmlNode2);
		if (text.Length > 0)
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
			_hasMinorGridlines = Conversions.ToBoolean(text);
		}
		text = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_BTW_CATEGORIES);
		if (text.Length > 0)
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
			_axisBetweenCategories = Conversions.ToBoolean(text);
		}
		text = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_OTHER_AXIS_CROSSES_AT);
		if (text.Length > 0)
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
			_otherAxisCrossesAt = Conversions.ToInteger(text);
		}
		_majorGridlinesLine = new Line(xmlNode);
		_minorGridlinesLine = new Line(xmlNode2);
		_axisLine = new Line(nd);
		XmlNode xmlNode3 = nd;
		_title = new AxisTitle(xmlNode3.SelectSingleNode(FormatConstants.NODE_TITLE), units);
		_displayUnitLabel = new DisplayUnitLabel(xmlNode3.SelectSingleNode(FormatConstants.NODE_DISPLAY_UNIT_LABEL), units);
		_tickMarks = new TickMarks(xmlNode3.SelectSingleNode(FormatConstants.NODE_TICK_MARKS));
		_tickLabels = new TickLabels(xmlNode3.SelectSingleNode(FormatConstants.NODE_TICK_LABELS));
		xmlNode3 = null;
		xmlNode = null;
		xmlNode2 = null;
	}

	internal override void A(Microsoft.Office.Interop.Excel.Chart A)
	{
		try
		{
			if (Operators.CompareString(_nodeName, FormatConstants.NODE_AXES_COMMON, TextCompare: false) == 0)
			{
				B(A);
				C(A);
				D(A);
				E(A);
			}
			else if (Operators.CompareString(_nodeName, FormatConstants.NODE_AXES_PRIMARY_CATEGORY, TextCompare: false) == 0)
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
				B(A);
			}
			else if (Operators.CompareString(_nodeName, FormatConstants.NODE_AXES_PRIMARY_VALUE, TextCompare: false) == 0)
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
				C(A);
			}
			else if (Operators.CompareString(_nodeName, FormatConstants.NODE_AXES_SECONDARY_CATEGORY, TextCompare: false) == 0)
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
				D(A);
			}
			else if (Operators.CompareString(_nodeName, FormatConstants.NODE_AXES_SECONDARY_VALUE, TextCompare: false) == 0)
			{
				E(A);
			}
			_ = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void B(Axis A)
	{
		Axis axis = A;
		if (_axisBetweenCategories.HasValue)
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
			try
			{
				axis.AxisBetweenCategories = _axisBetweenCategories.Value;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		if (_otherAxisCrossesAt.HasValue)
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
				axis.Crosses = (XlAxisCrosses)_otherAxisCrossesAt.Value;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		if (_displayUnit.HasValue)
		{
			try
			{
				axis.DisplayUnit = _displayUnit.Value;
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
			if (axis.DisplayUnit == (XlDisplayUnit)(-4114))
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					if (_displayUnitCustom.HasValue)
					{
						axis.DisplayUnitCustom = _displayUnitCustom.Value;
					}
					break;
				}
			}
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
		_axisLine.A(A.Format.Line);
		_title.A(A);
		_displayUnitLabel.A(A);
		_tickMarks.A(A);
		_tickLabels.A(A);
		D(A);
		E(A);
		axis = null;
	}

	private void B(Microsoft.Office.Interop.Excel.Chart A)
	{
		try
		{
			if (!Conversions.ToBoolean(((_Chart)A).get_HasAxis((object)XlAxisType.xlCategory, (object)XlAxisGroup.xlPrimary)))
			{
				return;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				_ = (Axis)A.Axes(XlAxisType.xlCategory);
				B((Axis)A.Axes(XlAxisType.xlCategory));
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void C(Microsoft.Office.Interop.Excel.Chart A)
	{
		try
		{
			if (!Conversions.ToBoolean(((_Chart)A).get_HasAxis((object)XlAxisType.xlValue, (object)XlAxisGroup.xlPrimary)))
			{
				return;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				B((Axis)A.Axes(XlAxisType.xlValue));
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void D(Microsoft.Office.Interop.Excel.Chart A)
	{
		try
		{
			if (Conversions.ToBoolean(((_Chart)A).get_HasAxis((object)XlAxisType.xlCategory, (object)XlAxisGroup.xlSecondary)))
			{
				B((Axis)A.Axes(XlAxisType.xlCategory, XlAxisGroup.xlSecondary));
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void E(Microsoft.Office.Interop.Excel.Chart A)
	{
		try
		{
			if (!Conversions.ToBoolean(((_Chart)A).get_HasAxis((object)XlAxisType.xlValue, (object)XlAxisGroup.xlSecondary)))
			{
				return;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				B((Axis)A.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary));
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void C(Axis A)
	{
		Axis axis = A;
		if (_axisBetweenCategories.HasValue)
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
			try
			{
				axis.AxisBetweenCategories = _axisBetweenCategories.Value;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		if (_displayUnit.HasValue)
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
				axis.DisplayUnit = _displayUnit.Value;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		try
		{
			if (axis.DisplayUnit == (XlDisplayUnit)(-4114))
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					if (!_displayUnitCustom.HasValue)
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
						axis.DisplayUnitCustom = _displayUnitCustom.Value;
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
		_axisLine.A(A.Format.Line);
		_title.A(A);
		_displayUnitLabel.A(A);
		_tickMarks.A(A);
		_tickLabels.A(A);
		D(A);
		E(A);
		axis = null;
	}

	private void D(Axis A)
	{
		try
		{
			Axis axis = A;
			if (_hasMajorGridlines.HasValue)
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
				axis.HasMajorGridlines = _hasMajorGridlines.Value;
			}
			if (axis.HasMajorGridlines)
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
				_majorGridlinesLine.A(axis.MajorGridlines.Format.Line);
			}
			axis = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void E(Axis A)
	{
		try
		{
			Axis axis = A;
			if (_hasMinorGridlines.HasValue)
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
				axis.HasMinorGridlines = _hasMinorGridlines.Value;
			}
			if (axis.HasMinorGridlines)
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
				_minorGridlinesLine.A(axis.MinorGridlines.Format.Line);
			}
			axis = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
