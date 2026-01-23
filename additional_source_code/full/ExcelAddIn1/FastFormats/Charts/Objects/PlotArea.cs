using System;
using System.Runtime.CompilerServices;
using System.Xml;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class PlotArea : ChartItem
{
	private new double? m_A;

	private new double? m_B;

	[CompilerGenerated]
	private new Fill m_A;

	[CompilerGenerated]
	private new Line m_A;

	[CompilerGenerated]
	private new double? m_C;

	[CompilerGenerated]
	private double? m_D;

	[CompilerGenerated]
	private double? m_E;

	[CompilerGenerated]
	private double? m_F;

	[CompilerGenerated]
	private new string m_A;

	private Fill _fill
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

	private Line _line
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

	private double? _insideTop
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

	private double? _insideLeft
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

	private double? _insideHeight
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

	private double? _insideWidth
	{
		[CompilerGenerated]
		get
		{
			return this.m_F;
		}
		[CompilerGenerated]
		set
		{
			this.m_F = value;
		}
	}

	private string _units
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

	public PlotArea(XmlNode nd, string units)
	{
		_insideTop = null;
		_insideLeft = null;
		_insideHeight = null;
		_insideWidth = null;
		_units = units;
		_fill = new Fill(nd);
		_line = new Line(nd);
		string attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_INSIDE_TOP);
		if (attributeValue.Length > 0)
		{
			_insideTop = modFunctionsConvert.CvtInvariantStrToDbl(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_INSIDE_LEFT);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			_insideLeft = modFunctionsConvert.CvtInvariantStrToDbl(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_INSIDE_HEIGHT);
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
			_insideHeight = modFunctionsConvert.CvtInvariantStrToDbl(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_INSIDE_WIDTH);
		if (attributeValue.Length <= 0)
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
			_insideWidth = modFunctionsConvert.CvtInvariantStrToDbl(attributeValue);
			return;
		}
	}

	internal override void A(Microsoft.Office.Interop.Excel.Chart A)
	{
		try
		{
			ChartFormat format = A.PlotArea.Format;
			_fill.A(format.Fill);
			_line.A(format.Line);
			format = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (!B())
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
			Microsoft.Office.Interop.Excel.Chart chart = A;
			try
			{
				C(chart.PlotArea);
				F(chart.PlotArea);
				B(chart.PlotArea);
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			chart = null;
			return;
		}
	}

	private void B(Microsoft.Office.Interop.Excel.PlotArea A)
	{
		Microsoft.Office.Interop.Excel.PlotArea plotArea = A;
		if (_insideTop.HasValue)
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
				plotArea.InsideTop = FormatUtil.GetDimensionInPoints(_insideTop.Value, _units);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		if (_insideLeft.HasValue)
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
				plotArea.InsideLeft = FormatUtil.GetDimensionInPoints(_insideLeft.Value, _units);
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		E(A);
		plotArea = null;
	}

	private void C(Microsoft.Office.Interop.Excel.PlotArea A)
	{
		D(A);
		Microsoft.Office.Interop.Excel.PlotArea plotArea = A;
		try
		{
			plotArea.InsideTop = 0.0;
			plotArea.InsideLeft = 0.0;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		plotArea = null;
	}

	private void D(Microsoft.Office.Interop.Excel.PlotArea A)
	{
		if (!_insideTop.HasValue)
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
			try
			{
				this.m_A = A.InsideTop;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		if (_insideLeft.HasValue)
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
			try
			{
				this.m_B = A.InsideLeft;
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

	private void E(Microsoft.Office.Interop.Excel.PlotArea A)
	{
		if (this.m_A.HasValue)
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
				A.InsideTop = this.m_A.Value;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		if (!this.m_B.HasValue)
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
				A.InsideLeft = this.m_B.Value;
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

	private bool B()
	{
		if (_insideTop.HasValue)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return true;
				}
			}
		}
		if (_insideLeft.HasValue)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		if (_insideHeight.HasValue)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		if (_insideWidth.HasValue)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		return false;
	}

	private void F(Microsoft.Office.Interop.Excel.PlotArea A)
	{
		Microsoft.Office.Interop.Excel.PlotArea plotArea = A;
		if (_insideHeight.HasValue)
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
			try
			{
				plotArea.InsideHeight = FormatUtil.GetDimensionInPoints(_insideHeight.Value, _units);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		if (_insideWidth.HasValue)
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
				plotArea.InsideWidth = FormatUtil.GetDimensionInPoints(_insideWidth.Value, _units);
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		plotArea = null;
	}
}
