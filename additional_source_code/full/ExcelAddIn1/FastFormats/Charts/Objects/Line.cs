using System;
using System.Runtime.CompilerServices;
using System.Xml;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class Line : BaseItem
{
	[CompilerGenerated]
	private int? m_A;

	[CompilerGenerated]
	private MsoLineDashStyle? m_A;

	[CompilerGenerated]
	private MsoLineStyle? m_A;

	[CompilerGenerated]
	private float? m_A;

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

	private MsoLineDashStyle? _dashStyle
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

	private MsoLineStyle? _style
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

	private float? _weight
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

	public Line(XmlNode nd)
	{
		_color = null;
		_dashStyle = null;
		_style = null;
		_weight = null;
		string attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_LINE_COLOR);
		string attributeValue2 = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_LINE_STYLE);
		string attributeValue3 = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_LINE_DASH);
		string attributeValue4 = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_LINE_WEIGHT);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (Operators.CompareString(attributeValue, FormatConstants.TRANSPARENCY.ToString(), TextCompare: false) != 0)
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
				_color = clsColors.RGB2Ole(attributeValue);
			}
			else
			{
				_color = FormatConstants.TRANSPARENCY;
			}
		}
		if (attributeValue2.Length > 0)
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
			_style = (MsoLineStyle)Conversions.ToInteger(attributeValue2);
		}
		if (attributeValue3.Length > 0)
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
			_dashStyle = (MsoLineDashStyle)Conversions.ToInteger(attributeValue3);
		}
		if (attributeValue4.Length <= 0)
		{
			return;
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			_weight = Conversions.ToSingle(attributeValue4);
			return;
		}
	}

	internal void A(LineFormat A)
	{
		if (!_color.HasValue && !_style.HasValue)
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
			if (!_dashStyle.HasValue)
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
				if (!_weight.HasValue)
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
					break;
				}
			}
		}
		A.Visible = MsoTriState.msoTrue;
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
			try
			{
				if (_color.Value != FormatConstants.TRANSPARENCY)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						A.ForeColor.RGB = _color.Value;
						break;
					}
				}
				else
				{
					A.Visible = MsoTriState.msoFalse;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		if (_dashStyle.HasValue && _dashStyle.Value > (MsoLineDashStyle)0)
		{
			try
			{
				A.DashStyle = _dashStyle.Value;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		if (_weight.HasValue && _weight.Value > 0f)
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
				A.Weight = _weight.Value;
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
		}
		if (!_style.HasValue)
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
			if (_style.Value <= (MsoLineStyle)0)
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
				try
				{
					A.Style = _style.Value;
					return;
				}
				catch (Exception ex7)
				{
					ProjectData.SetProjectError(ex7);
					Exception ex8 = ex7;
					ProjectData.ClearProjectError();
					return;
				}
			}
		}
	}
}
