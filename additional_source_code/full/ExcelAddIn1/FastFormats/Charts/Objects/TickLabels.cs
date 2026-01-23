using System;
using System.Runtime.CompilerServices;
using System.Xml;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public sealed class TickLabels : AxisItem
{
	[CompilerGenerated]
	private new XlTickLabelPosition? m_A;

	[CompilerGenerated]
	private new Font m_A;

	[CompilerGenerated]
	private new bool? m_A;

	[CompilerGenerated]
	private new int? m_A;

	[CompilerGenerated]
	private new int? B;

	[CompilerGenerated]
	private new XlHAlign? m_A;

	private XlTickLabelPosition? _tickLabelPosition
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

	private bool? _multiLevel
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

	private int? _offset
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
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	private XlHAlign? _alignment
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

	public TickLabels(XmlNode nd)
	{
		_tickLabelPosition = null;
		_multiLevel = null;
		_offset = null;
		_customAngle = null;
		_alignment = null;
		string attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_POSITION);
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
			_tickLabelPosition = (XlTickLabelPosition)Conversions.ToInteger(attributeValue);
		}
		_font = new Font(nd);
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_MULTILEVEL);
		if (attributeValue.Length > 0)
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
			_multiLevel = Conversions.ToBoolean(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_OFFSET);
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
			_offset = Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_CUSTOM_ANGLE);
		if (attributeValue.Length > 0)
		{
			_customAngle = Conversions.ToInteger(attributeValue);
		}
		attributeValue = clsXml.GetAttributeValue(nd, FormatConstants.ATTR_ALIGN);
		if (attributeValue.Length <= 0)
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			_alignment = (XlHAlign)Conversions.ToInteger(attributeValue);
			return;
		}
	}

	internal override void A(Axis A)
	{
		Axis axis = A;
		if (_tickLabelPosition.HasValue)
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
			try
			{
				axis.TickLabelPosition = _tickLabelPosition.Value;
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
			if (axis.TickLabelPosition != XlTickLabelPosition.xlTickLabelPositionNone)
			{
				Microsoft.Office.Interop.Excel.TickLabels tickLabels = axis.TickLabels;
				try
				{
					_font.A(tickLabels.Font);
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				if (_multiLevel.HasValue)
				{
					try
					{
						tickLabels.MultiLevel = _multiLevel.Value;
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						ProjectData.ClearProjectError();
					}
				}
				if (_offset.HasValue)
				{
					try
					{
						tickLabels.Offset = _offset.Value;
					}
					catch (Exception ex7)
					{
						ProjectData.SetProjectError(ex7);
						Exception ex8 = ex7;
						ProjectData.ClearProjectError();
					}
				}
				if (_customAngle.HasValue)
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
						tickLabels.Orientation = (XlTickLabelOrientation)checked(_customAngle.Value * -1);
					}
					catch (Exception ex9)
					{
						ProjectData.SetProjectError(ex9);
						Exception ex10 = ex9;
						ProjectData.ClearProjectError();
					}
				}
				if (_alignment.HasValue)
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
						tickLabels.Alignment = (int)_alignment.Value;
					}
					catch (Exception ex11)
					{
						ProjectData.SetProjectError(ex11);
						Exception ex12 = ex11;
						ProjectData.ClearProjectError();
					}
				}
				tickLabels = null;
			}
		}
		catch (Exception ex13)
		{
			ProjectData.SetProjectError(ex13);
			Exception ex14 = ex13;
			ProjectData.ClearProjectError();
		}
		axis = null;
	}
}
