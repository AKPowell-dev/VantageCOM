using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;
using A;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.Explorer;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public abstract class SheetItem : BaseItem
{
	[CompilerGenerated]
	private WorkbookItem m_A;

	[CompilerGenerated]
	private object m_A;

	private ObservableCollection<ResultItem> m_A;

	private int m_A;

	internal WorkbookItem Parent
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

	internal object Sheet
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = RuntimeHelpers.GetObjectValue(value);
		}
	}

	public ObservableCollection<ResultItem> Children
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(124354));
		}
	}

	public int ResultsCount
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(124371));
		}
	}

	public SheetItem(object sh, Microsoft.Office.Interop.Excel.Workbook wb, WorkbookItem wbi, Geometry geo)
		: base(Conversions.ToString(NewLateBinding.LateGet(sh, null, VH.A(19019), new object[0], null, null, null)), geo)
	{
		Sheet = RuntimeHelpers.GetObjectValue(sh);
		base.Workbook = wb;
		Parent = wbi;
		if (sh is Worksheet)
		{
			Children = new ObservableCollection<ResultItem>();
			if (((Worksheet)sh).ProtectContents)
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
				((BaseItem)this).Icon = Props.Icons.GeoLock;
				((BaseItem)this).Icon.Freeze();
			}
		}
		((BaseItem)this).IndentLevel = ((wbi != null) ? 1 : 0);
		((BaseItem)this).IsExpanded = true;
		A((XlSheetVisibility)Conversions.ToInteger(NewLateBinding.LateGet(sh, null, VH.A(41367), new object[0], null, null, null)));
	}

	internal void A(XlSheetVisibility A)
	{
		double opacity = ((A == XlSheetVisibility.xlSheetVisible) ? 1.0 : ((BaseItem)this).HIDDEN_OPACITY);
		if (Sheet is Worksheet)
		{
			base.IconColor = base.FontColor;
		}
		else
		{
			base.IconColor = Constants.ColorPalette.Green;
		}
		base.FontColor.Opacity = opacity;
		base.IconColor.Opacity = opacity;
	}

	internal void A(ResultItem A)
	{
		Children.Remove(A);
	}

	internal void A()
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Microsoft.Office.Interop.Excel.Workbook workbook = default(Microsoft.Office.Interop.Excel.Workbook);
		int count = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 468:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_0019;
						case 4:
							goto IL_002a;
						case 5:
							goto IL_0044;
						case 7:
							goto IL_005c;
						case 8:
							goto IL_008c;
						case 9:
							goto IL_00cb;
						case 10:
							goto IL_00f7;
						case 11:
							goto IL_0108;
						case 12:
							goto IL_0132;
						case 13:
							goto IL_0143;
						case 14:
							goto IL_0161;
						case 6:
						case 15:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 16:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_00cb:
					num2 = 9;
					NewLateBinding.LateSetComplex(Sheet, null, VH.A(41367), new object[1] { XlSheetVisibility.xlSheetHidden }, null, null, OptimisticSet: false, RValueBase: true);
					goto IL_00f7;
					IL_0007:
					num2 = 2;
					workbook = Parent.Workbook;
					goto IL_0019;
					IL_0019:
					num2 = 3;
					count = workbook.Sheets.Count;
					goto IL_002a;
					IL_002a:
					num2 = 4;
					if (count == 1)
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
						goto IL_0044;
					}
					goto IL_005c;
					IL_00f7:
					num2 = 10;
					workbook.Application.DisplayAlerts = false;
					goto IL_0108;
					IL_0108:
					num2 = 11;
					NewLateBinding.LateCall(Sheet, null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
					goto IL_0132;
					IL_0132:
					num2 = 12;
					workbook.Application.DisplayAlerts = true;
					goto IL_0143;
					IL_0044:
					num2 = 5;
					Forms.WarningMessage(VH.A(124396));
					break;
					IL_005c:
					num2 = 7;
					if (System.Windows.Forms.MessageBox.Show(VH.A(124507), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
					{
						break;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						break;
					}
					goto IL_008c;
					IL_0143:
					num2 = 13;
					if (workbook.Sheets.Count >= count)
					{
						break;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						break;
					}
					goto IL_0161;
					IL_008c:
					num2 = 8;
					if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(Sheet, null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetVeryHidden, TextCompare: false))
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
						goto IL_00cb;
					}
					goto IL_00f7;
					IL_0161:
					num2 = 14;
					Parent.Sheets.Remove(this);
					break;
					end_IL_0000_2:
					break;
				}
				num2 = 15;
				workbook = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 468;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	internal void B()
	{
		SendKeys.Send(VH.A(124594));
		System.Windows.Forms.Application.DoEvents();
		NewLateBinding.LateCall(NewLateBinding.LateGet(NewLateBinding.LateGet(Sheet, null, VH.A(124603), new object[0], null, null, null), null, VH.A(124626), new object[0], null, null, null), null, VH.A(124649), new object[1] { VH.A(124670) }, null, null, null, IgnoreReturn: true);
		System.Windows.Forms.Application.DoEvents();
	}

	internal void A(SheetItem A)
	{
		_ = Parent.Workbook.Application;
		Microsoft.Office.Interop.Excel.Workbook workbook = A.Parent.Workbook;
		object objectValue = RuntimeHelpers.GetObjectValue(A.Sheet);
		Microsoft.Office.Interop.Excel.Workbook workbook2 = Parent.Workbook;
		object objectValue2 = RuntimeHelpers.GetObjectValue(Sheet);
		object objectValue3;
		SheetItem sheetItem = default(SheetItem);
		if (!Workbooks.IsShared(workbook2, true, (System.Windows.Window)null))
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
			if (objectValue2 == null)
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
				objectValue2 = RuntimeHelpers.GetObjectValue(workbook2.Sheets[1]);
			}
			int num = Conversions.ToInteger(NewLateBinding.LateGet(objectValue2, null, VH.A(48135), new object[0], null, null, null));
			if (Operators.CompareString(workbook2.Name, workbook.Name, TextCompare: false) != 0)
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
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = workbook2.Sheets.GetEnumerator();
					while (enumerator.MoveNext())
					{
						if (!Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(RuntimeHelpers.GetObjectValue(enumerator.Current), null, VH.A(19019), new object[0], null, null, null), NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null), TextCompare: false))
						{
							continue;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							Forms.WarningMessage(VH.A(124693));
							throw new Exception();
						}
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_0175;
						}
						continue;
						end_IL_0175:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			object[] array;
			bool[] array2;
			NewLateBinding.LateCall(objectValue, null, VH.A(224), array = new object[1] { objectValue2 }, new string[1] { VH.A(51175) }, null, array2 = new bool[1] { true }, IgnoreReturn: true);
			if (array2[0])
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
				objectValue2 = RuntimeHelpers.GetObjectValue(array[0]);
			}
			objectValue3 = RuntimeHelpers.GetObjectValue(workbook.Application.ActiveSheet);
			if (A is WorksheetItem)
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
				sheetItem = new WorksheetItem(Parent, (Worksheet)objectValue3, (Microsoft.Office.Interop.Excel.Workbook)NewLateBinding.LateGet(objectValue3, null, VH.A(8701), new object[0], null, null, null), ((WorksheetItem)A).ResultsCount);
			}
			((BaseItem)sheetItem).IsSelected = true;
			Parent.Sheets.Insert(checked(num - 1), sheetItem);
		}
		workbook = null;
		workbook2 = null;
		objectValue = null;
		objectValue2 = null;
		objectValue3 = null;
		sheetItem = null;
	}

	internal void C()
	{
		if (Parent == null)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			int num = Parent.Sheets.IndexOf(this);
			if (num <= -1)
			{
				return;
			}
			int num2 = Conversions.ToInteger(Operators.SubtractObject(NewLateBinding.LateGet(Sheet, null, VH.A(48135), new object[0], null, null, null), 1));
			if (num == num2)
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
				Parent.Sheets.Move(num, num2);
				return;
			}
		}
	}

	internal void D()
	{
		Microsoft.Office.Interop.Excel.Sheets sheets = base.Workbook.Sheets;
		bool flag = false;
		try
		{
			int num = Conversions.ToInteger(Operators.SubtractObject(NewLateBinding.LateGet(Sheet, null, VH.A(48135), new object[0], null, null, null), 1));
			while (true)
			{
				if (num >= 1)
				{
					if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(sheets[num], null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetVisible, TextCompare: false))
					{
						A(RuntimeHelpers.GetObjectValue(sheets[num]));
						flag = true;
						break;
					}
					num = checked(num + -1);
					continue;
				}
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
				break;
			}
			if (!flag)
			{
				B(RuntimeHelpers.GetObjectValue(sheets[sheets.Count]));
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		sheets = null;
		C();
	}

	internal void E()
	{
		Microsoft.Office.Interop.Excel.Sheets sheets = base.Workbook.Sheets;
		bool flag = false;
		try
		{
			int num = Conversions.ToInteger(Operators.AddObject(NewLateBinding.LateGet(Sheet, null, VH.A(48135), new object[0], null, null, null), 1));
			int count = sheets.Count;
			for (int i = num; i <= count; i = checked(i + 1))
			{
				if (!Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(sheets[i], null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetVisible, TextCompare: false))
				{
					continue;
				}
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
				B(RuntimeHelpers.GetObjectValue(sheets[i]));
				flag = true;
				break;
			}
			if (!flag)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					A(RuntimeHelpers.GetObjectValue(sheets[1]));
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		sheets = null;
		C();
	}

	private void A(object A)
	{
		object sheet = Sheet;
		string memberName = VH.A(124835);
		object[] obj = new object[1] { A };
		object[] array = obj;
		string[] argumentNames = new string[1] { VH.A(51175) };
		bool[] obj2 = new bool[1] { true };
		bool[] array2 = obj2;
		NewLateBinding.LateCall(sheet, null, memberName, obj, argumentNames, null, obj2, IgnoreReturn: true);
		if (!array2[0])
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
			A = RuntimeHelpers.GetObjectValue(array[0]);
			return;
		}
	}

	private void B(object A)
	{
		object sheet = Sheet;
		string memberName = VH.A(124835);
		object[] obj = new object[1] { A };
		object[] array = obj;
		string[] argumentNames = new string[1] { VH.A(80163) };
		bool[] obj2 = new bool[1] { true };
		bool[] array2 = obj2;
		NewLateBinding.LateCall(sheet, null, memberName, obj, argumentNames, null, obj2, IgnoreReturn: true);
		if (array2[0])
		{
			A = RuntimeHelpers.GetObjectValue(array[0]);
		}
	}

	internal void F()
	{
		B(XlSheetVisibility.xlSheetVisible);
		NewLateBinding.LateCall(Sheet, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
	}

	internal void G()
	{
		B(XlSheetVisibility.xlSheetHidden);
	}

	internal void H()
	{
		B(XlSheetVisibility.xlSheetVeryHidden);
	}

	private void B(XlSheetVisibility A)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		application.EnableEvents = false;
		try
		{
			NewLateBinding.LateSetComplex(Sheet, null, VH.A(41367), new object[1] { A }, null, null, OptimisticSet: false, RValueBase: true);
			this.A(A);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application.EnableEvents = true;
		application = null;
	}
}
