using System;
using System.Runtime.CompilerServices;
using System.Xml;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class Font : BaseItem
{
	[CompilerGenerated]
	private double? m_A;

	[CompilerGenerated]
	private string m_A;

	[CompilerGenerated]
	private int? m_A;

	[CompilerGenerated]
	private bool? m_A;

	[CompilerGenerated]
	private new bool? m_B;

	private double? _size
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

	private string _name
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

	private int? _color
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

	private bool? _bold
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

	private bool? _italic
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

	public Font(XmlNode nd)
	{
		_size = null;
		_name = null;
		_color = null;
		_bold = null;
		_italic = null;
		string attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_FONT_COLOR);
		string attributeValue2 = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_FONT_NAME);
		string attributeValue3 = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_FONT_SIZE);
		string attributeValue4 = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_FONT_BOLD);
		string attributeValue5 = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_FONT_ITALIC);
		if (attributeValue.Length > 0)
		{
			_color = clsColors.RGB2Ole(attributeValue);
		}
		if (attributeValue2.Length > 0)
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
			_name = attributeValue2;
		}
		if (attributeValue3.Length > 0)
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
			_size = modFunctionsConvert.CvtInvariantStrToDbl(attributeValue3);
		}
		if (attributeValue4.Length > 0)
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
			_bold = Conversions.ToBoolean(attributeValue4);
		}
		if (attributeValue5.Length > 0)
		{
			_italic = Conversions.ToBoolean(attributeValue5);
		}
	}

	internal void A(Microsoft.Office.Interop.Excel.Chart A)
	{
		Microsoft.Office.Interop.Excel.Chart chart = A;
		try
		{
			this.A(chart.ChartArea.Font);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			if (Conversions.ToBoolean(((_Chart)chart).get_HasAxis((object)XlAxisType.xlValue, (object)XlAxisGroup.xlPrimary)))
			{
				B((Axis)chart.Axes(XlAxisType.xlValue));
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
			if (Conversions.ToBoolean(((_Chart)chart).get_HasAxis((object)XlAxisType.xlCategory, (object)XlAxisGroup.xlPrimary)))
			{
				B((Axis)chart.Axes(XlAxisType.xlCategory));
			}
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
		try
		{
			if (Conversions.ToBoolean(((_Chart)chart).get_HasAxis((object)XlAxisType.xlValue, (object)XlAxisGroup.xlSecondary)))
			{
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
					B((Axis)chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary));
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
		try
		{
			if (Conversions.ToBoolean(((_Chart)chart).get_HasAxis((object)XlAxisType.xlCategory, (object)XlAxisGroup.xlSecondary)))
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					B((Axis)chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlSecondary));
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
		try
		{
			if (chart.HasTitle)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					this.A(chart.ChartTitle.Font);
					break;
				}
			}
		}
		catch (Exception ex11)
		{
			ProjectData.SetProjectError(ex11);
			Exception ex12 = ex11;
			ProjectData.ClearProjectError();
		}
		try
		{
			if (chart.HasLegend)
			{
				this.A(chart.Legend.Font);
			}
		}
		catch (Exception ex13)
		{
			ProjectData.SetProjectError(ex13);
			Exception ex14 = ex13;
			ProjectData.ClearProjectError();
		}
		try
		{
			if (chart.HasDataTable)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					this.A(chart.DataTable.Font);
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
		chart = null;
	}

	private void B(Axis A)
	{
		Axis axis = A;
		try
		{
			if (axis.HasTitle)
			{
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
					this.A(axis.AxisTitle.Font);
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			if (axis.TickLabelPosition != XlTickLabelPosition.xlTickLabelPositionNone)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					this.A(axis.TickLabels.Font);
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
		axis = null;
	}

	internal void A(Microsoft.Office.Interop.Excel.Font A)
	{
		if (_color.HasValue)
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
			try
			{
				A.Color = _color.Value;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		if (_name != null)
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
				A.Name = _name;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		if (_size.HasValue)
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
				A.Size = _size.Value;
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
		}
		if (_bold.HasValue)
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
				A.Bold = _bold.Value;
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
		}
		if (!_italic.HasValue)
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
			try
			{
				A.Italic = _italic.Value;
				return;
			}
			catch (Exception ex9)
			{
				ProjectData.SetProjectError(ex9);
				Exception ex10 = ex9;
				ProjectData.ClearProjectError();
				return;
			}
		}
	}
}
