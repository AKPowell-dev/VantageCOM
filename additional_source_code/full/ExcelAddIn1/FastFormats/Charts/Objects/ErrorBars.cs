using System;
using System.Runtime.CompilerServices;
using System.Xml;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class ErrorBars : SeriesItem
{
	[CompilerGenerated]
	private new Line m_A;

	[CompilerGenerated]
	private new XlEndStyleCap? m_A;

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

	private XlEndStyleCap? _endStyle
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

	public ErrorBars(XmlNode nd)
	{
		_endStyle = null;
		_line = new Line(nd);
		string attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_END_STYLE);
		if (attributeValue.Length > 0)
		{
			_endStyle = (XlEndStyleCap)Conversions.ToInteger(attributeValue);
		}
	}

	internal override void A(Microsoft.Office.Interop.Excel.Series A)
	{
		Microsoft.Office.Interop.Excel.Series series = A;
		_line.A(series.ErrorBars.Format.Line);
		if (_endStyle.HasValue)
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
				series.ErrorBars.EndStyle = _endStyle.Value;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		series = null;
	}
}
