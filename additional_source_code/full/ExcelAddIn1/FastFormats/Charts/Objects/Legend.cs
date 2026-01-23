using System;
using System.Runtime.CompilerServices;
using System.Xml;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class Legend : ChartItem
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
	private new XlLegendPosition? m_A;

	[CompilerGenerated]
	private new string m_A;

	private bool? _hasLegend
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

	private XlLegendPosition? _position
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

	public Legend(XmlNode nd, string units)
	{
		_includeInLayout = null;
		_position = null;
		_units = units;
		_font = new Font(nd);
		_fill = new Fill(nd);
		_line = new Line(nd);
		string text = ((BaseItem)this).B(nd);
		if (text.Length > 0)
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
			_hasLegend = Conversions.ToBoolean(text);
		}
		text = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_POSITION);
		if (text.Length <= 0)
		{
			return;
		}
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			_position = (XlLegendPosition)Conversions.ToInteger(text);
			return;
		}
	}

	internal override void A(Microsoft.Office.Interop.Excel.Chart A)
	{
		if (_hasLegend.HasValue)
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
			try
			{
				A.HasLegend = _hasLegend.Value;
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
			if (!A.HasLegend)
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
				Microsoft.Office.Interop.Excel.Legend legend = A.Legend;
				try
				{
					_font.A(legend.Font);
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				try
				{
					_fill.A(legend.Format.Fill);
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
				try
				{
					_line.A(legend.Format.Line);
				}
				catch (Exception ex7)
				{
					ProjectData.SetProjectError(ex7);
					Exception ex8 = ex7;
					ProjectData.ClearProjectError();
				}
				if (_includeInLayout.HasValue)
				{
					try
					{
						legend.IncludeInLayout = _includeInLayout.Value;
					}
					catch (Exception ex9)
					{
						ProjectData.SetProjectError(ex9);
						Exception ex10 = ex9;
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
					try
					{
						legend.Position = _position.Value;
					}
					catch (Exception ex11)
					{
						ProjectData.SetProjectError(ex11);
						Exception ex12 = ex11;
						ProjectData.ClearProjectError();
					}
				}
				legend = null;
				return;
			}
		}
		catch (Exception ex13)
		{
			ProjectData.SetProjectError(ex13);
			Exception ex14 = ex13;
			ProjectData.ClearProjectError();
		}
	}
}
