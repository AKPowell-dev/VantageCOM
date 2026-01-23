using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros;
using MacabacusMacros.Explorer;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class PivotTableItem : ExploreItem
{
	private bool A;

	[CompilerGenerated]
	private PivotTable A;

	public override bool IsSelected
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(21693));
			Refresh();
		}
	}

	internal PivotTable PivotTable
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	public PivotTableItem(WorksheetItem wsi, PivotTable pvt)
		: base(wsi, Constants.ColorPalette.Orange.Clone(), Props.Icons.GeoPivot, 18)
	{
		PivotTable = pvt;
		base.Range = pvt.TableRange1;
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
			0 => VH.A(117433), 
			2 => VH.A(117456), 
			1 => VH.A(117485), 
			_ => VH.A(117433), 
		} + VH.A(17350) + PivotTable.TableRange1.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
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
				switch (5)
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
			isHighlighted = ((Operators.CompareString(strQuery, VH.A(117534), TextCompare: false) == 0) ? 1 : 0);
		}
		else
		{
			isHighlighted = 1;
		}
		((BaseItem)this).IsHighlighted = (byte)isHighlighted != 0;
	}

	public void RefreshTable()
	{
		PivotTable pivotTable = PivotTable;
		pivotTable.Application.EnableEvents = false;
		try
		{
			pivotTable.RefreshTable();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		pivotTable.Application.EnableEvents = true;
		pivotTable = null;
	}

	public void Clear()
	{
		try
		{
			PivotTable.ClearTable();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
	}
}
