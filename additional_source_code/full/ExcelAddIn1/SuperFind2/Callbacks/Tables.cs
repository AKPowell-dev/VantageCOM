using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.SuperFind2.Results;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class Tables
{
	internal static void A(WorksheetItem A, object B)
	{
		if (B is Microsoft.Office.Interop.Excel.Worksheet)
		{
			IEnumerator enumerator = ((Microsoft.Office.Interop.Excel.Worksheet)B).ListObjects.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					ListObject a = (ListObject)enumerator.Current;
					A.A(a);
				}
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
						return;
					}
				}
			}
			finally
			{
				IDisposable disposable = enumerator as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
		}
		Application application = MH.A.Application;
		IEnumerator enumerator2 = default(IEnumerator);
		try
		{
			enumerator2 = ((Range)B).Worksheet.ListObjects.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				ListObject listObject = (ListObject)enumerator2.Current;
				if (application.Intersect((Range)B, listObject.Range, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == null)
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
				A.A(listObject);
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_0213;
				}
				continue;
				end_IL_0213:
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
		application = null;
	}

	internal static void B(WorksheetItem A, object B)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Application application = default(Application);
		ListObject listObject = default(ListObject);
		IEnumerator enumerator = default(IEnumerator);
		Microsoft.Office.Interop.Excel.Worksheet worksheet = default(Microsoft.Office.Interop.Excel.Worksheet);
		QueryTables queryTables = default(QueryTables);
		IEnumerator enumerator2 = default(IEnumerator);
		QueryTable a = default(QueryTable);
		QueryTable queryTable = default(QueryTable);
		IEnumerator enumerator3 = default(IEnumerator);
		ListObject listObject2 = default(ListObject);
		IEnumerator enumerator4 = default(IEnumerator);
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
				case 1525:
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
							goto IL_0027;
						case 4:
							goto IL_0030;
						case 5:
							goto IL_003c;
						case 6:
							goto IL_0054;
						case 7:
							goto IL_0071;
						case 8:
							goto IL_007b;
						case 9:
							goto IL_0092;
						case 10:
							goto IL_00aa;
						case 11:
							goto IL_00ce;
						case 12:
							goto IL_00e6;
						case 13:
							goto IL_00f8;
						case 14:
							goto IL_0110;
						case 16:
							goto IL_0130;
						case 17:
							goto IL_0141;
						case 18:
							goto IL_0150;
						case 19:
							goto IL_015d;
						case 20:
							goto IL_0179;
						case 21:
							goto IL_019c;
						case 22:
							goto IL_030b;
						case 23:
							goto IL_0316;
						case 24:
							goto IL_0331;
						case 25:
							goto IL_0353;
						case 26:
							goto IL_037a;
						case 27:
							goto IL_0393;
						case 28:
							goto IL_04fc;
						case 29:
							goto IL_050c;
						case 30:
							goto IL_0525;
						case 31:
							goto IL_0547;
						case 15:
						case 32:
							goto IL_054d;
						case 33:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 34:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0393:
					num2 = 27;
					if (application.Intersect((Range)B, listObject.Range, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
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
						goto IL_04fc;
					}
					goto IL_050c;
					IL_0007:
					num2 = 2;
					if (B is Microsoft.Office.Interop.Excel.Worksheet)
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
						goto IL_0027;
					}
					goto IL_0130;
					IL_0331:
					num2 = 24;
					if (enumerator is IDisposable)
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
						(enumerator as IDisposable).Dispose();
					}
					goto IL_0353;
					IL_0547:
					num2 = 31;
					application = null;
					goto IL_054d;
					IL_054d:
					num2 = 32;
					worksheet = null;
					break;
					IL_0027:
					num2 = 3;
					worksheet = (Microsoft.Office.Interop.Excel.Worksheet)B;
					goto IL_0030;
					IL_0030:
					num2 = 4;
					queryTables = worksheet.QueryTables;
					goto IL_003c;
					IL_003c:
					num2 = 5;
					if (queryTables.Count > 0)
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
						goto IL_0054;
					}
					goto IL_00aa;
					IL_050c:
					num2 = 29;
					goto IL_050f;
					IL_0054:
					num2 = 6;
					enumerator2 = queryTables.GetEnumerator();
					goto IL_007d;
					IL_007d:
					if (enumerator2.MoveNext())
					{
						a = (QueryTable)enumerator2.Current;
						goto IL_0071;
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
					goto IL_0092;
					IL_019c:
					num2 = 21;
					if (application.Intersect((Range)B, queryTable.ResultRange, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
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
						goto IL_030b;
					}
					goto IL_0316;
					IL_0092:
					num2 = 9;
					if (enumerator2 is IDisposable)
					{
						(enumerator2 as IDisposable).Dispose();
					}
					goto IL_00aa;
					IL_030b:
					num2 = 22;
					A.A(queryTable);
					goto IL_0316;
					IL_04fc:
					num2 = 28;
					A.A(listObject.QueryTable);
					goto IL_050c;
					IL_0071:
					num2 = 7;
					A.A(a);
					goto IL_007b;
					IL_007b:
					num2 = 8;
					goto IL_007d;
					IL_00aa:
					num2 = 10;
					enumerator3 = worksheet.ListObjects.GetEnumerator();
					goto IL_00fb;
					IL_00fb:
					if (enumerator3.MoveNext())
					{
						listObject2 = (ListObject)enumerator3.Current;
						goto IL_00ce;
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
					goto IL_0110;
					IL_0316:
					num2 = 23;
					goto IL_0319;
					IL_0110:
					num2 = 14;
					if (enumerator3 is IDisposable)
					{
						(enumerator3 as IDisposable).Dispose();
					}
					goto IL_054d;
					IL_050f:
					if (enumerator4.MoveNext())
					{
						listObject = (ListObject)enumerator4.Current;
						goto IL_037a;
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
					goto IL_0525;
					IL_0353:
					num2 = 25;
					enumerator4 = worksheet.ListObjects.GetEnumerator();
					goto IL_050f;
					IL_00ce:
					num2 = 11;
					if (listObject2.QueryTable != null)
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
						goto IL_00e6;
					}
					goto IL_00f8;
					IL_037a:
					num2 = 26;
					if (listObject.QueryTable != null)
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
						goto IL_0393;
					}
					goto IL_050c;
					IL_00e6:
					num2 = 12;
					A.A(listObject2.QueryTable);
					goto IL_00f8;
					IL_00f8:
					num2 = 13;
					goto IL_00fb;
					IL_0130:
					num2 = 16;
					application = MH.A.Application;
					goto IL_0141;
					IL_0141:
					num2 = 17;
					worksheet = ((Range)B).Worksheet;
					goto IL_0150;
					IL_0150:
					num2 = 18;
					queryTables = worksheet.QueryTables;
					goto IL_015d;
					IL_015d:
					num2 = 19;
					if (queryTables.Count > 0)
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
						goto IL_0179;
					}
					goto IL_0353;
					IL_0525:
					num2 = 30;
					if (enumerator4 is IDisposable)
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
						(enumerator4 as IDisposable).Dispose();
					}
					goto IL_0547;
					IL_0179:
					num2 = 20;
					enumerator = queryTables.GetEnumerator();
					goto IL_0319;
					IL_0319:
					if (enumerator.MoveNext())
					{
						queryTable = (QueryTable)enumerator.Current;
						goto IL_019c;
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
					goto IL_0331;
					end_IL_0000_2:
					break;
				}
				num2 = 33;
				queryTables = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 1525;
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

	internal static void C(WorksheetItem A, object B)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		PivotTables pivotTables = default(PivotTables);
		PivotTable pivotTable = default(PivotTable);
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		PivotTable a = default(PivotTable);
		Application application = default(Application);
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
				case 828:
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
							goto IL_0027;
						case 4:
							goto IL_0048;
						case 5:
							goto IL_0062;
						case 6:
							goto IL_0080;
						case 7:
							goto IL_008a;
						case 8:
							goto IL_009f;
						case 10:
							goto IL_00c8;
						case 11:
							goto IL_00d9;
						case 12:
							goto IL_0100;
						case 13:
							goto IL_011b;
						case 14:
							goto IL_013d;
						case 15:
							goto IL_029a;
						case 16:
							goto IL_02a5;
						case 17:
							goto IL_02b4;
						case 18:
							goto IL_02cc;
						case 9:
						case 19:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 20:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_00d9:
					num2 = 11;
					pivotTables = (PivotTables)((Range)B).Worksheet.PivotTables(RuntimeHelpers.GetObjectValue(Missing.Value));
					goto IL_0100;
					IL_0007:
					num2 = 2;
					if (B is Microsoft.Office.Interop.Excel.Worksheet)
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						goto IL_0027;
					}
					goto IL_00c8;
					IL_0100:
					num2 = 12;
					if (pivotTables.Count > 0)
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
						goto IL_011b;
					}
					goto IL_02cc;
					IL_029a:
					num2 = 15;
					A.A(pivotTable);
					goto IL_02a5;
					IL_011b:
					num2 = 13;
					enumerator = pivotTables.GetEnumerator();
					goto IL_02a8;
					IL_0027:
					num2 = 3;
					pivotTables = (PivotTables)((Microsoft.Office.Interop.Excel.Worksheet)B).PivotTables(RuntimeHelpers.GetObjectValue(Missing.Value));
					goto IL_0048;
					IL_0048:
					num2 = 4;
					if (pivotTables.Count <= 0)
					{
						break;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						break;
					}
					goto IL_0062;
					IL_02a8:
					if (enumerator.MoveNext())
					{
						pivotTable = (PivotTable)enumerator.Current;
						goto IL_013d;
					}
					goto IL_02b4;
					IL_0062:
					num2 = 5;
					enumerator2 = pivotTables.GetEnumerator();
					goto IL_008c;
					IL_008c:
					if (enumerator2.MoveNext())
					{
						a = (PivotTable)enumerator2.Current;
						goto IL_0080;
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
					goto IL_009f;
					IL_02b4:
					num2 = 17;
					if (enumerator is IDisposable)
					{
						(enumerator as IDisposable).Dispose();
					}
					goto IL_02cc;
					IL_009f:
					num2 = 8;
					if (!(enumerator2 is IDisposable))
					{
						break;
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
					(enumerator2 as IDisposable).Dispose();
					break;
					IL_02a5:
					num2 = 16;
					goto IL_02a8;
					IL_013d:
					num2 = 14;
					if (application.Intersect((Range)B, pivotTable.TableRange1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
					{
						goto IL_029a;
					}
					goto IL_02a5;
					IL_02cc:
					num2 = 18;
					application = null;
					break;
					IL_0080:
					num2 = 6;
					A.A(a);
					goto IL_008a;
					IL_008a:
					num2 = 7;
					goto IL_008c;
					IL_00c8:
					num2 = 10;
					application = MH.A.Application;
					goto IL_00d9;
					end_IL_0000_2:
					break;
				}
				num2 = 19;
				pivotTables = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 828;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
	}

	internal static void D(WorksheetItem A, object B)
	{
		if (B is Microsoft.Office.Interop.Excel.Worksheet)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)B;
					if (worksheet.AutoFilterMode)
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
						A.A(worksheet.AutoFilter);
					}
					worksheet = null;
					return;
				}
				}
			}
		}
		Microsoft.Office.Interop.Excel.Worksheet worksheet2 = ((Range)B).Worksheet;
		if (worksheet2.AutoFilterMode)
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
			if (worksheet2.Application.Intersect(worksheet2.AutoFilter.Range, (Range)B, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
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
				A.A(worksheet2.AutoFilter);
			}
		}
		worksheet2 = null;
	}
}
