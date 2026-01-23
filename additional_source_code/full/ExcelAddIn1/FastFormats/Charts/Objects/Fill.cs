using System;
using System.Runtime.CompilerServices;
using System.Xml;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class Fill : BaseItem
{
	[CompilerGenerated]
	private int? m_A;

	[CompilerGenerated]
	private new int? B;

	private int? _foreColor
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

	private int? _backColor
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

	public Fill(XmlNode nd)
	{
		_foreColor = null;
		_backColor = null;
		string attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_FILL_FORE_COLOR);
		if (attributeValue.Length > 0)
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
			if (Operators.CompareString(attributeValue, FormatConstants.TRANSPARENCY.ToString(), TextCompare: false) != 0)
			{
				_foreColor = clsColors.RGB2Ole(attributeValue);
			}
			else
			{
				_foreColor = FormatConstants.TRANSPARENCY;
			}
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_FILL_BACK_COLOR);
		if (attributeValue.Length <= 0)
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
			if (Operators.CompareString(attributeValue, FormatConstants.TRANSPARENCY.ToString(), TextCompare: false) != 0)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						_backColor = clsColors.RGB2Ole(attributeValue);
						return;
					}
				}
			}
			_backColor = FormatConstants.TRANSPARENCY;
			return;
		}
	}

	internal void A(Microsoft.Office.Interop.Excel.FillFormat A)
	{
		if (!_foreColor.HasValue)
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
			if (!_backColor.HasValue)
			{
				return;
			}
		}
		try
		{
			if (_foreColor.Value != FormatConstants.TRANSPARENCY)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						A.ForeColor.RGB = _foreColor.Value;
						A.Visible = MsoTriState.msoTrue;
						A.Solid();
						return;
					}
				}
			}
			A.Visible = MsoTriState.msoFalse;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
