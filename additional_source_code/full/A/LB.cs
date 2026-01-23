using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

internal sealed class LB
{
	internal static bool A(List<int> A)
	{
		bool result;
		if (A == null)
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
			result = false;
		}
		else
		{
			try
			{
				Sheets worksheets = MH.A.Application.ActiveWorkbook.Worksheets;
				List<string> list = new List<string>();
				using (List<int>.Enumerator enumerator = A.GetEnumerator())
				{
					while (enumerator.MoveNext())
					{
						int current = enumerator.Current;
						try
						{
							if (!(worksheets.get_Item((object)current) is Worksheet worksheet))
							{
								continue;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								list.Add(worksheet.Name);
								break;
							}
						}
						catch (Exception projectError)
						{
							ProjectData.SetProjectError(projectError);
							ProjectData.ClearProjectError();
						}
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_009b;
						}
						continue;
						end_IL_009b:
						break;
					}
				}
				if (list.Count == 0)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						result = false;
						break;
					}
				}
				else
				{
					((Sheets)worksheets.get_Item((object)list.ToArray())).Select(RuntimeHelpers.GetObjectValue(Missing.Value));
					result = true;
				}
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				result = false;
				ProjectData.ClearProjectError();
			}
			finally
			{
				Worksheet worksheet2 = null;
				Sheets worksheets = null;
			}
		}
		return result;
	}

	internal static List<int> A()
	{
		List<int> list = new List<int>();
		try
		{
			Sheets selectedSheets = MH.A.Application.ActiveWindow.SelectedSheets;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = selectedSheets.GetEnumerator();
				while (enumerator.MoveNext())
				{
					object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
					if (objectValue is Worksheet)
					{
						list.Add(((Worksheet)objectValue).Index);
					}
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
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (7)
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
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		finally
		{
			object objectValue = null;
			Sheets selectedSheets = null;
		}
		return list;
	}
}
