using System;
using System.Runtime.CompilerServices;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class UpDownBars : BaseItem
{
	[CompilerGenerated]
	private Fill m_A;

	[CompilerGenerated]
	private Line m_A;

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

	public UpDownBars(XmlNode nd)
	{
		_fill = new Fill(nd);
		_line = new Line(nd);
	}

	internal void A(ChartFormat A)
	{
		ChartFormat chartFormat = A;
		try
		{
			_fill.A(chartFormat.Fill);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			_line.A(chartFormat.Line);
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		chartFormat = null;
	}
}
