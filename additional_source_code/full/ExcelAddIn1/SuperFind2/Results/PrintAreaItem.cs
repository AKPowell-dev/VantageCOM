using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class PrintAreaItem : ExploreItem
{
	private bool A;

	public override bool IsSelected
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(21693));
			Refresh();
		}
	}

	public PrintAreaItem(WorksheetItem wsi, Range rng)
		: base(wsi, Constants.ColorPalette.Blue.Clone(), Props.Icons.GeoPrinter, 2)
	{
		base.Range = rng;
		Refresh();
	}

	public override void Refresh()
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_0005: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000a: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Expected I4, but got Unknown
		Language applicationLanguage = clsEnvironment.ApplicationLanguage;
		((BaseItem)this).Label = (applicationLanguage - 1) switch
		{
			0 => VH.A(117547), 
			2 => VH.A(117568), 
			1 => VH.A(117603), 
			_ => VH.A(117547), 
		} + VH.A(17350) + base.Range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
	}

	public override void Delete()
	{
		if (MessageBox.Show(VH.A(117638), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Worksheet worksheet = base.Worksheet;
			if (Operators.CompareString(worksheet.PageSetup.PrintArea, "", TextCompare: false) != 0)
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
				List<string> list = new List<string>();
				string listSeparator = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
				string[] array = Strings.Split(worksheet.PageSetup.PrintArea, listSeparator, -1, CompareMethod.Text);
				foreach (string cell in array)
				{
					Range range = ((_Worksheet)worksheet).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value));
					if (Operators.CompareString(range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), base.Range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) != 0)
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
						list.Add(range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
					}
					range = null;
				}
				worksheet.PageSetup.PrintArea = Strings.Join(list.ToArray(), listSeparator);
				list = null;
			}
			worksheet = null;
			base.Parent.A(this);
			return;
		}
	}

	public override void Search(string strQuery)
	{
		((BaseItem)this).IsHighlighted = ((BaseItem)this).Label.ToLower().Contains(strQuery) || Operators.CompareString(strQuery, VH.A(117735), TextCompare: false) == 0;
	}
}
