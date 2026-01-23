using System;
using System.Runtime.CompilerServices;
using System.Xml;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class DisplayUnitLabel : AxisItem
{
	[CompilerGenerated]
	private new bool? m_A;

	[CompilerGenerated]
	private new Fill m_A;

	[CompilerGenerated]
	private new Font m_A;

	[CompilerGenerated]
	private new Line m_A;

	[CompilerGenerated]
	private new XlHAlign? m_A;

	[CompilerGenerated]
	private new XlVAlign? m_A;

	[CompilerGenerated]
	private new int? m_A;

	[CompilerGenerated]
	private new XlChartElementPosition? m_A;

	[CompilerGenerated]
	private new double? m_A;

	[CompilerGenerated]
	private new double? B;

	[CompilerGenerated]
	private new string m_A;

	private bool? _hasLabel
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

	private Font _font
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

	private XlHAlign? _horizontalAlignment
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

	private XlVAlign? _verticalAlignment
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

	private int? _customAngle
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

	private XlChartElementPosition? _position
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

	private double? _top
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

	private double? _left
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

	public DisplayUnitLabel(XmlNode nd, string units)
	{
		_hasLabel = null;
		_horizontalAlignment = null;
		_verticalAlignment = null;
		_customAngle = null;
		_position = null;
		_top = null;
		_left = null;
		_units = units;
		_fill = new Fill(nd);
		_line = new Line(nd);
		_font = new Font(nd);
		string text = B(nd);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			_hasLabel = Conversions.ToBoolean(text);
		}
		text = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_ALIGN_HORIZ);
		if (text.Length > 0)
		{
			_horizontalAlignment = (XlHAlign)Conversions.ToInteger(text);
		}
		text = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_ALIGN_VERT);
		if (text.Length > 0)
		{
			_verticalAlignment = (XlVAlign)Conversions.ToInteger(text);
		}
		text = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_CUSTOM_ANGLE);
		if (text.Length > 0)
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
			_customAngle = Conversions.ToInteger(text);
		}
		text = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_POSITION);
		if (text.Length > 0)
		{
			_position = (XlChartElementPosition)Conversions.ToInteger(text);
		}
		text = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_TOP);
		if (text.Length > 0)
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
			_top = modFunctionsConvert.CvtInvariantStrToDbl(text);
		}
		text = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_LEFT);
		if (text.Length <= 0)
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
			_left = modFunctionsConvert.CvtInvariantStrToDbl(text);
			return;
		}
	}

	internal override void A(Axis A)
	{
		try
		{
			Axis axis = A;
			if (_hasLabel.HasValue)
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
					axis.HasDisplayUnitLabel = _hasLabel.Value;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			try
			{
				if (axis.HasDisplayUnitLabel)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						Microsoft.Office.Interop.Excel.DisplayUnitLabel displayUnitLabel = axis.DisplayUnitLabel;
						try
						{
							_font.A(displayUnitLabel.Font);
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
						}
						try
						{
							_fill.A(displayUnitLabel.Format.Fill);
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
							ProjectData.ClearProjectError();
						}
						try
						{
							_line.A(displayUnitLabel.Format.Line);
						}
						catch (Exception ex7)
						{
							ProjectData.SetProjectError(ex7);
							Exception ex8 = ex7;
							ProjectData.ClearProjectError();
						}
						if (_horizontalAlignment.HasValue)
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
								displayUnitLabel.HorizontalAlignment = _horizontalAlignment.Value;
							}
							catch (Exception ex9)
							{
								ProjectData.SetProjectError(ex9);
								Exception ex10 = ex9;
								ProjectData.ClearProjectError();
							}
						}
						if (_verticalAlignment.HasValue)
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
								displayUnitLabel.VerticalAlignment = _verticalAlignment.Value;
							}
							catch (Exception ex11)
							{
								ProjectData.SetProjectError(ex11);
								Exception ex12 = ex11;
								ProjectData.ClearProjectError();
							}
						}
						if (_customAngle.HasValue)
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
								displayUnitLabel.Orientation = checked(_customAngle.Value * -1);
							}
							catch (Exception ex13)
							{
								ProjectData.SetProjectError(ex13);
								Exception ex14 = ex13;
								ProjectData.ClearProjectError();
							}
						}
						if (_position.HasValue)
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
							if (_position.Value == XlChartElementPosition.xlChartElementPositionAutomatic)
							{
								displayUnitLabel.Position = _position.Value;
							}
							else
							{
								try
								{
									if (_top.HasValue)
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
										displayUnitLabel.Top = FormatUtil.GetDimensionInPoints(_top.Value, _units);
									}
									if (_left.HasValue)
									{
										while (true)
										{
											switch (6)
											{
											case 0:
												continue;
											}
											displayUnitLabel.Left = FormatUtil.GetDimensionInPoints(_left.Value, _units);
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
							}
						}
						displayUnitLabel = null;
						break;
					}
				}
			}
			catch (Exception ex17)
			{
				ProjectData.SetProjectError(ex17);
				Exception ex18 = ex17;
				ProjectData.ClearProjectError();
			}
			axis = null;
		}
		catch (Exception ex19)
		{
			ProjectData.SetProjectError(ex19);
			Exception ex20 = ex19;
			ProjectData.ClearProjectError();
		}
	}
}
