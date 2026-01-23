using System;
using System.Runtime.CompilerServices;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class ChartArea : ChartItem
{
	[CompilerGenerated]
	private new Fill m_A;

	[CompilerGenerated]
	private new Line m_A;

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

	public ChartArea(XmlNode nd)
	{
		_fill = new Fill(nd);
		_line = new Line(nd);
	}

	internal override void A(Microsoft.Office.Interop.Excel.Chart A)
	{
		ChartFormat format = A.ChartArea.Format;
		try
		{
			_fill.A(format.Fill);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			_line.A(format.Line);
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		format = null;
		_ = null;
	}
}
