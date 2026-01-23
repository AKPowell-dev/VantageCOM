using System;
using System.Runtime.CompilerServices;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class LeaderLines : SeriesItem
{
	[CompilerGenerated]
	private new Line m_A;

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

	public LeaderLines(XmlNode nd)
	{
		_line = new Line(nd);
	}

	internal override void A(Microsoft.Office.Interop.Excel.Series A)
	{
		try
		{
			_line.A(A.LeaderLines.Format.Line);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
