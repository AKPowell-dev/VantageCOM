using System;
using System.Runtime.CompilerServices;
using System.Xml;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class DataTable : ChartItem
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
	private bool? D;

	[CompilerGenerated]
	private bool? E;

	private bool? _hasDataTable
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

	private bool? _showLegendKey
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

	private bool? _hasBorderHorizontal
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

	private bool? _hasBorderVertical
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[CompilerGenerated]
		set
		{
			D = value;
		}
	}

	private bool? _hasBorderOutline
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[CompilerGenerated]
		set
		{
			E = value;
		}
	}

	public DataTable(XmlNode nd)
	{
		_showLegendKey = null;
		_hasBorderHorizontal = null;
		_hasBorderVertical = null;
		_hasBorderOutline = null;
		_font = new Font(nd);
		_fill = new Fill(nd);
		_line = new Line(nd);
		string text = ((BaseItem)this).B(nd);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			_hasDataTable = Conversions.ToBoolean(text);
		}
		text = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_SHOW_LEGEND_KEY);
		if (text.Length > 0)
		{
			_showLegendKey = Conversions.ToBoolean(text);
		}
		text = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_HAS_BORDER_HORIZ);
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
			_hasBorderHorizontal = Conversions.ToBoolean(text);
		}
		text = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_HAS_BORDER_VERT);
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
			_hasBorderVertical = Conversions.ToBoolean(text);
		}
		text = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_HAS_BORDER_OUTLINE);
		if (text.Length <= 0)
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			_hasBorderOutline = Conversions.ToBoolean(text);
			return;
		}
	}

	internal override void A(Microsoft.Office.Interop.Excel.Chart A)
	{
		if (_hasDataTable.HasValue)
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
			try
			{
				A.HasDataTable = _hasDataTable.Value;
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
			if (!A.HasDataTable)
			{
				return;
			}
			Microsoft.Office.Interop.Excel.DataTable dataTable = A.DataTable;
			try
			{
				_font.A(dataTable.Font);
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			try
			{
				_fill.A(dataTable.Format.Fill);
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			try
			{
				_line.A(dataTable.Format.Line);
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
			if (_showLegendKey.HasValue)
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
					dataTable.ShowLegendKey = _showLegendKey.Value;
				}
				catch (Exception ex9)
				{
					ProjectData.SetProjectError(ex9);
					Exception ex10 = ex9;
					ProjectData.ClearProjectError();
				}
			}
			if (_hasBorderHorizontal.HasValue)
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
				try
				{
					dataTable.HasBorderHorizontal = _hasBorderHorizontal.Value;
				}
				catch (Exception ex11)
				{
					ProjectData.SetProjectError(ex11);
					Exception ex12 = ex11;
					ProjectData.ClearProjectError();
				}
			}
			if (_hasBorderOutline.HasValue)
			{
				try
				{
					dataTable.HasBorderOutline = _hasBorderOutline.Value;
				}
				catch (Exception ex13)
				{
					ProjectData.SetProjectError(ex13);
					Exception ex14 = ex13;
					ProjectData.ClearProjectError();
				}
			}
			if (_hasBorderVertical.HasValue)
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
				try
				{
					dataTable.HasBorderVertical = _hasBorderVertical.Value;
				}
				catch (Exception ex15)
				{
					ProjectData.SetProjectError(ex15);
					Exception ex16 = ex15;
					ProjectData.ClearProjectError();
				}
			}
			dataTable = null;
		}
		catch (Exception ex17)
		{
			ProjectData.SetProjectError(ex17);
			Exception ex18 = ex17;
			ProjectData.ClearProjectError();
		}
	}
}
