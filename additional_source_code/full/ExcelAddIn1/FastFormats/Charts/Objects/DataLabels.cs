using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Xml;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class DataLabels : SeriesItem
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<KeyValuePair<string, string>, string> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal string A(KeyValuePair<string, string> A)
		{
			return A.Value;
		}
	}

	[CompilerGenerated]
	internal sealed class WF
	{
		public string A;

		[SpecialName]
		internal bool A(KeyValuePair<string, string> A)
		{
			return Operators.CompareString(A.Key.Trim(), this.A.Trim(), TextCompare: false) == 0;
		}
	}

	[CompilerGenerated]
	private new Font m_A;

	[CompilerGenerated]
	private new Fill m_A;

	[CompilerGenerated]
	private new Line m_A;

	[CompilerGenerated]
	private new bool? m_A;

	[CompilerGenerated]
	private new bool? m_B;

	[CompilerGenerated]
	private bool? C;

	[CompilerGenerated]
	private bool? D;

	[CompilerGenerated]
	private bool? E;

	[CompilerGenerated]
	private bool? F;

	[CompilerGenerated]
	private new string m_A;

	[CompilerGenerated]
	private bool? G;

	[CompilerGenerated]
	private new XlHAlign? m_A;

	[CompilerGenerated]
	private new int? m_A;

	[CompilerGenerated]
	private new XlDataLabelPosition? m_A;

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

	private bool? _autoText
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
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	private bool? _showCategoryName
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

	private bool? _showBubbleSize
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

	private bool? _showPercentage
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

	private bool? _showSeriesName
	{
		[CompilerGenerated]
		get
		{
			return F;
		}
		[CompilerGenerated]
		set
		{
			F = value;
		}
	}

	private string _separator
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

	private bool? _showValue
	{
		[CompilerGenerated]
		get
		{
			return G;
		}
		[CompilerGenerated]
		set
		{
			G = value;
		}
	}

	private XlHAlign? _horizAlign
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

	private int? _customAngle
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

	private XlDataLabelPosition? _position
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

	public DataLabels(XmlNode nd)
	{
		_autoText = null;
		_showLegendKey = null;
		_showCategoryName = null;
		_showBubbleSize = null;
		_showPercentage = null;
		_showSeriesName = null;
		_separator = null;
		_showValue = null;
		_horizAlign = null;
		_customAngle = null;
		_position = null;
		_font = new Font(nd);
		_fill = new Fill(nd);
		_line = new Line(nd);
		string attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_AUTO_TEXT);
		if (attributeValue.Length > 0)
		{
			_autoText = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_SHOW_LEGEND_KEY);
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
			_showLegendKey = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_SHOW_CAT_NAME);
		if (attributeValue.Length > 0)
		{
			_showCategoryName = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_SEPARATOR);
		if (attributeValue.Length > 0)
		{
			_separator = B(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_SHOW_BUBBLE_SIZE);
		if (attributeValue.Length > 0)
		{
			_showBubbleSize = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_SHOW_PCT);
		if (attributeValue.Length > 0)
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
			_showPercentage = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_SHOW_SERIES_NAME);
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
			_showSeriesName = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_SHOW_VALUE);
		if (attributeValue.Length > 0)
		{
			_showValue = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_ALIGN_HORIZ);
		if (attributeValue.Length > 0)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			_horizAlign = (XlHAlign)Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_CUSTOM_ANGLE);
		if (attributeValue.Length > 0)
		{
			_customAngle = Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_POSITION);
		if (attributeValue.Length <= 0)
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
			_position = (XlDataLabelPosition)Conversions.ToInteger(attributeValue);
			return;
		}
	}

	internal override void A(Microsoft.Office.Interop.Excel.Series A)
	{
		Microsoft.Office.Interop.Excel.DataLabels dataLabels = (Microsoft.Office.Interop.Excel.DataLabels)A.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
		try
		{
			_font.A(dataLabels.Font);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			_fill.A(dataLabels.Format.Fill);
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		try
		{
			_line.A(dataLabels.Format.Line);
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
		if (_autoText.HasValue)
		{
			try
			{
				dataLabels.AutoText = _autoText.Value;
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
		}
		if (_separator != null)
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
			try
			{
				if (object.Equals(modFunctionsConvert.CvtToInt((object)_separator), 1))
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						dataLabels.Separator = XlDataLabelSeparator.xlDataLabelSeparatorDefault;
						break;
					}
				}
				else if (object.Equals(_separator, '\r'.ToString()))
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						dataLabels.Separator = '\r'.ToString();
						break;
					}
				}
				else
				{
					dataLabels.Separator = _separator;
				}
			}
			catch (Exception ex9)
			{
				ProjectData.SetProjectError(ex9);
				Exception ex10 = ex9;
				ProjectData.ClearProjectError();
			}
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
				dataLabels.ShowLegendKey = _showLegendKey.Value;
			}
			catch (Exception ex11)
			{
				ProjectData.SetProjectError(ex11);
				Exception ex12 = ex11;
				ProjectData.ClearProjectError();
			}
		}
		if (_showCategoryName.HasValue)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			try
			{
				dataLabels.ShowCategoryName = _showCategoryName.Value;
			}
			catch (Exception ex13)
			{
				ProjectData.SetProjectError(ex13);
				Exception ex14 = ex13;
				ProjectData.ClearProjectError();
			}
		}
		if (_showBubbleSize.HasValue)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			try
			{
				dataLabels.ShowBubbleSize = _showBubbleSize.Value;
			}
			catch (Exception ex15)
			{
				ProjectData.SetProjectError(ex15);
				Exception ex16 = ex15;
				ProjectData.ClearProjectError();
			}
		}
		if (_showPercentage.HasValue)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			try
			{
				dataLabels.ShowPercentage = _showPercentage.Value;
			}
			catch (Exception ex17)
			{
				ProjectData.SetProjectError(ex17);
				Exception ex18 = ex17;
				ProjectData.ClearProjectError();
			}
		}
		if (_showSeriesName.HasValue)
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
				dataLabels.ShowSeriesName = _showSeriesName.Value;
			}
			catch (Exception ex19)
			{
				ProjectData.SetProjectError(ex19);
				Exception ex20 = ex19;
				ProjectData.ClearProjectError();
			}
		}
		if (_showValue.HasValue)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			try
			{
				dataLabels.ShowValue = _showValue.Value;
			}
			catch (Exception ex21)
			{
				ProjectData.SetProjectError(ex21);
				Exception ex22 = ex21;
				ProjectData.ClearProjectError();
			}
		}
		if (_horizAlign.HasValue)
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
				dataLabels.HorizontalAlignment = _horizAlign.Value;
			}
			catch (Exception ex23)
			{
				ProjectData.SetProjectError(ex23);
				Exception ex24 = ex23;
				ProjectData.ClearProjectError();
			}
		}
		if (_customAngle.HasValue)
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
				dataLabels.Orientation = checked(_customAngle.Value * -1);
			}
			catch (Exception ex25)
			{
				ProjectData.SetProjectError(ex25);
				Exception ex26 = ex25;
				ProjectData.ClearProjectError();
			}
		}
		if (_position.HasValue)
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
			try
			{
				dataLabels.Position = _position.Value;
			}
			catch (Exception ex27)
			{
				ProjectData.SetProjectError(ex27);
				Exception ex28 = ex27;
				ProjectData.ClearProjectError();
			}
		}
		dataLabels = null;
	}

	private string B(string A)
	{
		IEnumerable<KeyValuePair<string, string>> source = from keyValuePair in FormatUtil.GetSeparatorDictionary()
			where Operators.CompareString(keyValuePair.Key.Trim(), A.Trim(), TextCompare: false) == 0
			select keyValuePair;
		Func<KeyValuePair<string, string>, string> selector;
		if (_Closure_0024__.A == null)
		{
			selector = (_Closure_0024__.A = [SpecialName] (KeyValuePair<string, string> keyValuePair) => keyValuePair.Value);
		}
		else
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
			selector = _Closure_0024__.A;
		}
		string text = source.Select(selector).FirstOrDefault();
		if (!string.IsNullOrEmpty(text))
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					return text;
				}
			}
		}
		return A;
	}
}
