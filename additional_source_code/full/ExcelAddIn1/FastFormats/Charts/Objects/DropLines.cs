using System;
using System.Runtime.CompilerServices;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class DropLines : BaseItem
{
	[CompilerGenerated]
	private Line m_A;

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

	public DropLines(XmlNode nd)
	{
		_line = new Line(nd);
	}

	internal void A(Microsoft.Office.Interop.Excel.DropLines A)
	{
		try
		{
			_line.A(A.Format.Line);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
