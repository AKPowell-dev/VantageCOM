using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using A;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros.Explorer;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class ErrorItem : ExploreItem
{
	private bool m_A;

	private Visibility m_A;

	[CompilerGenerated]
	private XlErrorChecks m_A;

	public override bool IsSelected
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(21693));
			Refresh();
		}
	}

	public Visibility ErrorWrapVisibility
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(118729));
		}
	}

	private XlErrorChecks ErrorType
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

	public ErrorItem(WorksheetItem wsi, Range rng, XlErrorChecks chk)
		: base(wsi, Constants.ColorPalette.Red.Clone(), Props.Icons.GeoError, 25)
	{
		base.Range = rng;
		ErrorType = chk;
		ErrorWrapVisibility = ((!Operators.ConditionalCompareObjectEqual(rng.Cells.CountLarge, 1, TextCompare: false)) ? Visibility.Collapsed : Visibility.Visible);
		Refresh();
	}

	public override void Refresh()
	{
		((BaseItem)this).Label = A() + VH.A(17350) + base.Range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
	}

	public override void Delete()
	{
		throw new NotImplementedException();
	}

	public override void Search(string strQuery)
	{
		int isHighlighted;
		if (!((BaseItem)this).Label.ToLower().Contains(strQuery))
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
			isHighlighted = ((Operators.CompareString(strQuery, VH.A(118716), TextCompare: false) == 0) ? 1 : 0);
		}
		else
		{
			isHighlighted = 1;
		}
		((BaseItem)this).IsHighlighted = (byte)isHighlighted != 0;
	}

	public void ErrorWrap()
	{
		if (ErrorType != XlErrorChecks.xlEvaluateToError)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					throw new NotImplementedException();
				}
			}
		}
		bool flag = false;
		string text;
		if (KH.A.ErrorValuePrompt)
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
			text = Forms.InputBox(VH.A(118768), VH.A(118803), VH.A(118884));
			if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
		}
		else
		{
			text = KH.A.DefaultErrorValue;
			if (Operators.CompareString(text, null, TextCompare: false) == 0)
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
				text = "";
			}
		}
		if (!Versioned.IsNumeric(text))
		{
			text = VH.A(39830) + text + VH.A(39830);
			flag = true;
		}
		Range range = base.Range;
		if (Conversions.ToBoolean(range.HasFormula))
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
			if (Strings.InStr(Conversions.ToString(range.Formula), VH.A(118889)) == 0)
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
				string text2 = Regex.Match(Conversions.ToString(range.Formula), VH.A(118904)).ToString();
				string text3 = Regex.Replace(Conversions.ToString(range.Formula), VH.A(118904), "");
				range.Formula = text2 + VH.A(79125) + text3 + VH.A(2378) + text + VH.A(39904);
				if (flag)
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
					range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;
				}
			}
		}
		range = null;
		base.Parent.A(this);
	}

	private string A()
	{
		switch (ErrorType)
		{
		case XlErrorChecks.xlEvaluateToError:
			base.Tooltip = VH.A(118915);
			return VH.A(118988);
		case XlErrorChecks.xlEmptyCellReferences:
			base.Tooltip = VH.A(119015);
			return VH.A(119120);
		case XlErrorChecks.xlInconsistentFormula:
			base.Tooltip = VH.A(119161);
			return VH.A(119270);
		case XlErrorChecks.xlInconsistentListFormula:
			base.Tooltip = VH.A(119311);
			return VH.A(119416);
		case XlErrorChecks.xlListDataValidation:
			base.Tooltip = VH.A(119467);
			return VH.A(119597);
		case XlErrorChecks.xlNumberAsText:
			base.Tooltip = VH.A(119652);
			return VH.A(119735);
		case XlErrorChecks.xlOmittedCells:
			base.Tooltip = VH.A(119778);
			return VH.A(119891);
		case XlErrorChecks.xlTextDate:
			base.Tooltip = VH.A(119930);
			return VH.A(120027);
		case XlErrorChecks.xlUnlockedFormulaCells:
			base.Tooltip = VH.A(120082);
			return VH.A(120177);
		default:
			return "";
		}
	}
}
