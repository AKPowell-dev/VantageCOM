using System;
using System.Runtime.CompilerServices;
using System.Xml;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class TickMarks : AxisItem
{
	[CompilerGenerated]
	private new XlTickMark? m_A;

	[CompilerGenerated]
	private new XlTickMark? B;

	private XlTickMark? _majorTickMark
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

	private XlTickMark? _minorTickMark
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

	public TickMarks(XmlNode nd)
	{
		_majorTickMark = null;
		_minorTickMark = null;
		string attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_MAJOR);
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
			_majorTickMark = (XlTickMark)Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_MINOR);
		if (attributeValue.Length <= 0)
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
			_minorTickMark = (XlTickMark)Conversions.ToInteger(attributeValue);
			return;
		}
	}

	internal override void A(Axis A)
	{
		try
		{
			Axis axis = A;
			if (_majorTickMark.HasValue)
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
					axis.MajorTickMark = _majorTickMark.Value;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			if (_minorTickMark.HasValue)
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
					axis.MinorTickMark = _minorTickMark.Value;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
			}
			axis = null;
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
	}
}
