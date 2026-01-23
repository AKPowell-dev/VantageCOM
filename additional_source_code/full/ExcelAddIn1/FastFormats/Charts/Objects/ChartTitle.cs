using System;
using System.Runtime.CompilerServices;
using System.Xml;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class ChartTitle : ChartItem
{
	[CompilerGenerated]
	private new bool? m_A;

	[CompilerGenerated]
	private new Font m_A;

	[CompilerGenerated]
	private new Fill m_A;

	[CompilerGenerated]
	private new Line m_A;

	[CompilerGenerated]
	private new bool? B;

	[CompilerGenerated]
	private new bool? C;

	[CompilerGenerated]
	private new string m_A;

	private bool? _hasTitle
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

	private bool? _includeInLayout
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

	private bool? _automatic
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[CompilerGenerated]
		set
		{
			C = value;
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

	public ChartTitle(XmlNode nd, string units)
	{
		_includeInLayout = null;
		_automatic = null;
		_units = units;
		_font = new Font(nd);
		_fill = new Fill(nd);
		_line = new Line(nd);
		string text = ((BaseItem)this).B(nd);
		if (text.Length > 0)
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
			_hasTitle = Conversions.ToBoolean(text);
		}
		text = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_AUTOMATIC);
		if (text.Length <= 0)
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
			_automatic = Conversions.ToBoolean(text);
			return;
		}
	}

	internal override void A(Microsoft.Office.Interop.Excel.Chart A)
	{
		if (_hasTitle.HasValue)
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
			A.HasTitle = _hasTitle.Value;
		}
		try
		{
			if (!A.HasTitle)
			{
				return;
			}
			Microsoft.Office.Interop.Excel.ChartTitle chartTitle = A.ChartTitle;
			try
			{
				_font.A(chartTitle.Font);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			try
			{
				_fill.A(chartTitle.Format.Fill);
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			try
			{
				_line.A(chartTitle.Format.Line);
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			try
			{
				if (_includeInLayout.HasValue)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						chartTitle.IncludeInLayout = _includeInLayout.Value;
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
				if (_automatic.HasValue)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						if (_automatic == true)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								chartTitle.Position = XlChartElementPosition.xlChartElementPositionAutomatic;
								break;
							}
						}
						else
						{
							chartTitle.Position = XlChartElementPosition.xlChartElementPositionCustom;
						}
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
			chartTitle = null;
		}
		catch (Exception ex11)
		{
			ProjectData.SetProjectError(ex11);
			Exception ex12 = ex11;
			ProjectData.ClearProjectError();
		}
	}
}
