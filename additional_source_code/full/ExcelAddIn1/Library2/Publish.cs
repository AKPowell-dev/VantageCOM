using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Library2;

public sealed class Publish
{
	public static bool Validate(Microsoft.Office.Interop.Excel.Workbook wb, string strErrorsMessage)
	{
		Microsoft.Office.Interop.Excel.Workbook workbook = wb;
		int num = 0;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = workbook.Styles.GetEnumerator();
			while (true)
			{
				if (enumerator.MoveNext())
				{
					if (((Style)enumerator.Current).BuiltIn)
					{
						continue;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					num = checked(num + 1);
					if (num <= 100)
					{
						continue;
					}
					if (MessageBox.Show(VH.A(85078), VH.A(40448), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_0072;
							}
							continue;
							end_IL_0072:
							break;
						}
						break;
					}
					return false;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_0090;
					}
					continue;
					end_IL_0090:
					break;
				}
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		try
		{
			Array array = (Array)workbook.LinkSources(XlLink.xlExcelLinks);
			if (array != null)
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
				if (array.Length > 0)
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
					if (MessageBox.Show(VH.A(85516), VH.A(40448), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) == DialogResult.No)
					{
						array = null;
						return false;
					}
				}
			}
			array = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (workbook.Names.Count > 10 && MessageBox.Show(VH.A(85812), VH.A(40448), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) == DialogResult.No)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					return false;
				}
			}
		}
		Range range = null;
		IEnumerator enumerator2 = default(IEnumerator);
		try
		{
			enumerator2 = workbook.Worksheets.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				object objectValue = RuntimeHelpers.GetObjectValue(enumerator2.Current);
				try
				{
					range = (Range)NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue, null, VH.A(82416), new object[0], null, null, null), null, VH.A(86222), new object[2]
					{
						XlCellType.xlCellTypeFormulas,
						XlSpecialCellsValue.xlErrors
					}, null, null, null);
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				if (range == null)
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
					if (MessageBox.Show(strErrorsMessage + VH.A(86247), VH.A(40448), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
					{
						range = null;
						return false;
					}
					break;
				}
				break;
			}
		}
		finally
		{
			if (enumerator2 is IDisposable)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					(enumerator2 as IDisposable).Dispose();
					break;
				}
			}
		}
		range = null;
		workbook = null;
		return true;
	}
}
