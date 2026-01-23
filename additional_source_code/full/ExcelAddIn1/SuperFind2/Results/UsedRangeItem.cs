using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Sheets;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class UsedRangeItem : ExploreItem
{
	private bool m_A;

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

	public UsedRangeItem(WorksheetItem wsi)
		: base(wsi, Constants.ColorPalette.Blue.Clone(), Props.Icons.GeoUsedRange, 1)
	{
		Refresh();
	}

	public override void Refresh()
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_0005: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected I4, but got Unknown
		Language applicationLanguage = clsEnvironment.ApplicationLanguage;
		string text = (applicationLanguage - 1) switch
		{
			0 => VH.A(113570), 
			2 => VH.A(122583), 
			1 => VH.A(122614), 
			3 => VH.A(122643), 
			_ => VH.A(113570), 
		};
		base.Range = base.Worksheet.UsedRange;
		((BaseItem)this).Label = text + VH.A(17350) + base.Range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
	}

	public override void Delete()
	{
		throw new NotImplementedException();
	}

	public override void Search(string strQuery)
	{
		((BaseItem)this).IsHighlighted = ((BaseItem)this).Label.ToLower().Contains(strQuery) || Operators.CompareString(strQuery, VH.A(122680), TextCompare: false) == 0;
	}

	internal void A()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num5 = default(int);
		int num = default(int);
		int num3 = default(int);
		Worksheet worksheet = default(Worksheet);
		int num6 = default(int);
		Microsoft.Office.Interop.Excel.Workbook workbook = default(Microsoft.Office.Interop.Excel.Workbook);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		DialogResult dialogResult = default(DialogResult);
		Worksheet worksheet2 = default(Worksheet);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				object instance;
				string memberName;
				object[] obj;
				bool[] obj2;
				object instance2;
				string memberName2;
				object[] obj3;
				object[] array;
				bool[] obj4;
				bool[] array2;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					num5 = 0;
					goto IL_0005;
				case 1022:
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
							goto IL_0005;
						case 3:
							goto IL_000a;
						case 4:
							goto IL_0024;
						case 5:
							goto IL_0044;
						case 6:
							goto IL_0068;
						case 7:
							goto IL_006f;
						case 9:
							goto IL_007a;
						case 8:
						case 10:
							goto IL_009f;
						case 11:
							goto IL_00a2;
						case 12:
							goto IL_00b7;
						case 13:
							goto IL_00be;
						case 14:
							goto IL_00d1;
						case 15:
							goto IL_00dc;
						case 16:
							goto IL_00e7;
						case 17:
							goto IL_00f2;
						case 18:
							goto IL_00fc;
						case 19:
							goto IL_0105;
						case 20:
							goto IL_010b;
						case 21:
							goto IL_0124;
						case 22:
							goto IL_0253;
						case 23:
							goto IL_0256;
						case 24:
							goto IL_0267;
						case 25:
							goto IL_026c;
						case 26:
							goto IL_0277;
						case 27:
							goto IL_0282;
						case 28:
							goto IL_028d;
						case 29:
							goto IL_0296;
						case 30:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 31:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0253:
					worksheet = null;
					goto IL_0256;
					IL_0005:
					num2 = 2;
					num6 = 0;
					goto IL_000a;
					IL_000a:
					num2 = 3;
					workbook = base.Parent.Parent.Workbook;
					goto IL_0024;
					IL_0024:
					num2 = 4;
					if (!workbook.Saved)
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
						goto IL_0044;
					}
					goto IL_007a;
					IL_0277:
					num2 = 26;
					application.DisplayAlerts = true;
					goto IL_0282;
					IL_0282:
					num2 = 27;
					application.EnableEvents = true;
					goto IL_028d;
					IL_026c:
					num2 = 25;
					application.ScreenUpdating = true;
					goto IL_0277;
					IL_0044:
					num2 = 5;
					dialogResult = MessageBox.Show(VH.A(122691), VH.A(40448), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
					goto IL_0068;
					IL_0068:
					num2 = 6;
					if (dialogResult == DialogResult.Yes)
					{
						goto IL_006f;
					}
					goto IL_009f;
					IL_006f:
					num2 = 7;
					workbook.Save();
					goto IL_009f;
					IL_007a:
					num2 = 9;
					dialogResult = MessageBox.Show(VH.A(123191), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
					goto IL_009f;
					IL_009f:
					workbook = null;
					goto IL_00a2;
					IL_00a2:
					num2 = 11;
					if (dialogResult == DialogResult.Cancel)
					{
						goto end_IL_0000_3;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						break;
					}
					goto IL_00b7;
					IL_028d:
					num2 = 28;
					Refresh();
					goto IL_0296;
					IL_00b7:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_00be;
					IL_00be:
					num2 = 13;
					application = MH.A.Application;
					goto IL_00d1;
					IL_00d1:
					num2 = 14;
					application.ScreenUpdating = false;
					goto IL_00dc;
					IL_00dc:
					num2 = 15;
					application.EnableEvents = false;
					goto IL_00e7;
					IL_00e7:
					num2 = 16;
					application.DisplayAlerts = false;
					goto IL_00f2;
					IL_00f2:
					num2 = 17;
					worksheet2 = base.Worksheet;
					goto IL_00fc;
					IL_00fc:
					num2 = 18;
					ExcelAddIn1.Sheets.Protection.Unprotect(worksheet2);
					goto IL_0105;
					IL_0105:
					num2 = 19;
					worksheet = worksheet2;
					goto IL_010b;
					IL_010b:
					num2 = 20;
					if (!worksheet.ProtectContents)
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
						goto IL_0124;
					}
					goto IL_0253;
					IL_0296:
					num2 = 29;
					instance = NewLateBinding.LateGet(base.Workbook, null, VH.A(123607), new object[0], null, null, null);
					memberName = VH.A(123653);
					obj = new object[2] { num5, num6 };
					array = obj;
					obj2 = new bool[2] { true, true };
					array2 = obj2;
					NewLateBinding.LateCall(instance, null, memberName, obj, null, null, obj2, IgnoreReturn: true);
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
						num5 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					if (!array2[1])
					{
						break;
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
					num6 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[1]), typeof(int));
					break;
					IL_0124:
					num2 = 21;
					instance2 = NewLateBinding.LateGet(base.Workbook, null, VH.A(123607), new object[0], null, null, null);
					memberName2 = VH.A(123624);
					obj3 = new object[5]
					{
						worksheet2,
						worksheet.Rows.CountLarge,
						worksheet.Columns.CountLarge,
						num5,
						num6
					};
					array = obj3;
					obj4 = new bool[5] { true, false, false, true, true };
					array2 = obj4;
					NewLateBinding.LateCall(instance2, null, memberName2, obj3, null, null, obj4, IgnoreReturn: true);
					if (array2[0])
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
						worksheet2 = (Worksheet)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(Worksheet));
					}
					if (array2[3])
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
						num5 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[3]), typeof(int));
					}
					if (array2[4])
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
						num6 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[4]), typeof(int));
					}
					goto IL_0253;
					IL_0256:
					num2 = 23;
					worksheet2.UsedRange.Select();
					goto IL_0267;
					IL_0267:
					num2 = 24;
					worksheet2 = null;
					goto IL_026c;
					end_IL_0000_2:
					break;
				}
				application = null;
				break;
				end_IL_0000:;
			}
			catch (object obj5) when (obj5 is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj5);
				try0000_dispatch = 1022;
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
			switch (2)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}
}
