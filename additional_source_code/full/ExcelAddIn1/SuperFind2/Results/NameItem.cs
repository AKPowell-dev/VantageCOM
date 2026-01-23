using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Media;
using A;
using ExcelAddIn1.Formulas;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros.Explorer;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class NameItem : ExploreItem
{
	[CompilerGenerated]
	internal sealed class MF
	{
		public string A;

		public MF(MF A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal bool A(Range A)
		{
			return Operators.CompareString(A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), this.A, TextCompare: false) == 0;
		}
	}

	private readonly Color m_A;

	private bool m_A;

	[CompilerGenerated]
	private Name m_A;

	[CompilerGenerated]
	private List<Range> m_A;

	[CompilerGenerated]
	private string m_A;

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

	internal Name Name
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

	private List<Range> NavigateRanges
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

	private string DialogMessage
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

	public NameItem(WorksheetItem wsi, Name nm, Range rng, bool blnBad, string strTooltip)
		: base(wsi, Constants.ColorPalette.LightGreen.Clone(), Props.Icons.GeoName, 35)
	{
		this.m_A = Colors.Firebrick;
		Name = nm;
		base.Range = rng;
		((BaseItem)this).Label = nm.Name + VH.A(17350) + rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		base.Tooltip = strTooltip;
		if (blnBad)
		{
			base.FontColor = new SolidColorBrush(this.m_A);
			base.IconColor = new SolidColorBrush(this.m_A);
		}
		if (!nm.Visible)
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
			base.FontColor.Opacity = ((BaseItem)this).HIDDEN_OPACITY;
			base.IconColor.Opacity = ((BaseItem)this).HIDDEN_OPACITY;
		}
		Refresh();
	}

	public override void Refresh()
	{
		try
		{
			base.Range = Name.RefersToRange;
			((BaseItem)this).Label = Name.Name + VH.A(17350) + base.Range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public override void Delete()
	{
		if (MessageBox.Show(VH.A(116583) + Name.Name + VH.A(43025), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
		{
			return;
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Name.Delete();
			base.Parent.A(this);
			return;
		}
	}

	public override void Search(string strQuery)
	{
		((BaseItem)this).IsHighlighted = ((BaseItem)this).Label.ToLower().Contains(strQuery) || Operators.CompareString(strQuery, VH.A(115495), TextCompare: false) == 0;
	}

	internal void A()
	{
		Name.Visible = false;
		base.FontColor.Opacity = ((BaseItem)this).HIDDEN_OPACITY;
		base.IconColor.Opacity = ((BaseItem)this).HIDDEN_OPACITY;
	}

	internal void B()
	{
		Name.Visible = true;
		base.FontColor.Opacity = 1.0;
		base.IconColor.Opacity = 1.0;
	}

	internal void C()
	{
		if (MessageBox.Show(VH.A(116668), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
		{
			return;
		}
		int count = default(int);
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Microsoft.Office.Interop.Excel.Application application = Name.Application;
			List<Range> list = new List<Range>();
			Name name = Name;
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			Range activeCell = application.ActiveCell;
			try
			{
				list = ExcelAddIn1.Formulas.Names.B(list, name);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			ExcelAddIn1.Formulas.Names.A(activeCell);
			application.ScreenUpdating = true;
			application.EnableEvents = true;
			try
			{
				count = list.Count;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			if (count == 0)
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
				Forms.InfoMessage(VH.A(116755));
			}
			else
			{
				NavigateRanges = list;
				string[] obj = new string[9]
				{
					VH.A(116911),
					name.Name,
					VH.A(116944),
					Conversions.ToString(count),
					VH.A(41385),
					(count == 1) ? VH.A(116966) : VH.A(116955),
					VH.A(116975),
					null,
					null
				};
				string text;
				if (count != 1)
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
					text = VH.A(117026);
				}
				else
				{
					text = VH.A(117049);
				}
				obj[7] = text;
				obj[8] = VH.A(117068);
				DialogMessage = string.Concat(obj);
				A(null, "");
			}
			activeCell = null;
			list = null;
			name = null;
			application = null;
			base.Parent.A(this);
			return;
		}
	}

	internal void D()
	{
		Microsoft.Office.Interop.Excel.Application application = Name.Application;
		Name name = Name;
		XlCalculation calculation = application.Calculation;
		List<Range> list = new List<Range>();
		List<Range> list2 = new List<Range>();
		try
		{
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			application.Calculation = XlCalculation.xlCalculationManual;
			Range activeCell = application.ActiveCell;
			Range refersToRange = default(Range);
			try
			{
				refersToRange = name.RefersToRange;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			list = ExcelAddIn1.Formulas.Names.GetDependents(refersToRange, name);
			int num = default(int);
			using (List<Range>.Enumerator enumerator = list.GetEnumerator())
			{
				MF mF = default(MF);
				while (enumerator.MoveNext())
				{
					Range current = enumerator.Current;
					if (Operators.CompareString(current.Worksheet.Name, name.RefersToRange.Worksheet.Name, TextCompare: false) != 0)
					{
						continue;
					}
					try
					{
						current.ApplyNames(new string[1] { name.Name }, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlApplyNamesOrder.xlRowThenColumn, RuntimeHelpers.GetObjectValue(Missing.Value));
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
					finally
					{
						num = Conversions.ToInteger(Operators.AddObject(num, current.Cells.CountLarge));
					}
					try
					{
						mF = new MF(mF);
						mF.A = current.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
						if (list2.Find(mF.A) != null)
						{
							continue;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							list2.Add(current);
							break;
						}
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_01e3;
					}
					continue;
					end_IL_01e3:
					break;
				}
			}
			try
			{
				ExcelAddIn1.Formulas.Names.A(activeCell);
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
			finally
			{
				application.Calculation = calculation;
				application.ScreenUpdating = true;
				application.EnableEvents = true;
			}
			try
			{
				string[] obj = new string[9]
				{
					VH.A(117071),
					name.Name,
					VH.A(116944),
					Conversions.ToString(num),
					VH.A(41385),
					null,
					null,
					null,
					null
				};
				string text;
				if (num != 1)
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
					text = VH.A(116955);
				}
				else
				{
					text = VH.A(116966);
				}
				obj[5] = text;
				obj[6] = VH.A(117100);
				string text2;
				if (num != 1)
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
					text2 = VH.A(117026);
				}
				else
				{
					text2 = VH.A(117049);
				}
				obj[7] = text2;
				obj[8] = VH.A(117068);
				string dialogMessage = string.Concat(obj);
				DialogMessage = dialogMessage;
				NavigateRanges = list2;
				if (!NavigateRanges.Any())
				{
					return;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					A(null, "");
					return;
				}
			}
			catch (Exception ex9)
			{
				ProjectData.SetProjectError(ex9);
				Exception ex10 = ex9;
				ProjectData.ClearProjectError();
			}
		}
		catch (Exception ex11)
		{
			ProjectData.SetProjectError(ex11);
			Exception ex12 = ex11;
			ProjectData.ClearProjectError();
		}
		finally
		{
			Range refersToRange = null;
			Range activeCell = null;
			list2 = null;
			list = null;
			name = null;
			application = null;
		}
	}

	private void A(List<Range> A = null, string B = "")
	{
		if (A == null)
		{
			A = NavigateRanges;
		}
		if (Operators.CompareString(B, "", TextCompare: false) == 0)
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
			B = DialogMessage;
		}
		if (MessageBox.Show(B, VH.A(40448), MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) != DialogResult.Yes)
		{
			return;
		}
		object instance = NewLateBinding.LateGet(base.Workbook, null, VH.A(117153), new object[0], null, null, null);
		string memberName = VH.A(117180);
		object[] obj = new object[1] { A };
		object[] array = obj;
		bool[] obj2 = new bool[1] { true };
		bool[] array2 = obj2;
		NewLateBinding.LateCall(instance, null, memberName, obj, null, null, obj2, IgnoreReturn: true);
		if (!array2[0])
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
			A = (List<Range>)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(List<Range>));
			return;
		}
	}
}
