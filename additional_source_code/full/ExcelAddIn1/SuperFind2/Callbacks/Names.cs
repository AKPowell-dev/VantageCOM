using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.Links;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class Names
{
	[CompilerGenerated]
	private static Dictionary<Microsoft.Office.Interop.Excel.Workbook, List<Name>> m_A;

	private static Dictionary<Microsoft.Office.Interop.Excel.Workbook, List<Name>> NamesDictionary
	{
		[CompilerGenerated]
		get
		{
			return Names.m_A;
		}
		[CompilerGenerated]
		set
		{
			Names.m_A = value;
		}
	}

	internal static List<Name> A(Microsoft.Office.Interop.Excel.Workbook A)
	{
		if (NamesDictionary != null)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (NamesDictionary.ContainsKey(A))
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						return NamesDictionary[A];
					}
				}
			}
		}
		List<Name> list = new List<Name>();
		bool flag = false;
		try
		{
			Microsoft.Office.Interop.Excel.Names names = A.Names;
			int count = names.Count;
			if (count > 100)
			{
				if (MessageBox.Show(VH.A(103971) + count + VH.A(104018), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
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
					flag = true;
				}
			}
			else if (count > 1000)
			{
				Forms.WarningMessage(VH.A(104460));
				flag = true;
			}
			if (!flag)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					int num = count;
					Name name;
					for (int i = 1; i <= num; name = null, i = checked(i + 1))
					{
						name = names.Item(i, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						if (name.Visible)
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
							if (!Names.IsNative(name.Name))
							{
								goto IL_014c;
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
						}
						if (!Names.A(name))
						{
							continue;
						}
						goto IL_014c;
						IL_014c:
						list.Add(name);
					}
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
		finally
		{
			Microsoft.Office.Interop.Excel.Names names = null;
		}
		if (NamesDictionary == null)
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
			NamesDictionary = new Dictionary<Microsoft.Office.Interop.Excel.Workbook, List<Name>>();
		}
		NamesDictionary.Add(A, list);
		return list;
	}

	internal static bool A(Name A)
	{
		bool result;
		try
		{
			int num;
			if (A.Name.StartsWith(Base.LINK_PREFIX))
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
				num = (Conversions.ToBoolean(Operators.CompareObjectNotEqual(A.RefersToRange.Cells.CountLarge, A.RefersToRange.Worksheet.Cells.CountLarge, TextCompare: false)) ? 1 : 0);
			}
			else
			{
				num = 0;
			}
			result = Conversions.ToBoolean((byte)num != 0);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}
}
