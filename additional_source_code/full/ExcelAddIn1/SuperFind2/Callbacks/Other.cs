using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Charts;
using ExcelAddIn1.SuperFind2.Results;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class Other
{
	internal static void A(WorksheetItem A, object B)
	{
		if (B is Microsoft.Office.Interop.Excel.Worksheet)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					{
						foreach (SparklineGroup sparklineGroup in ((Microsoft.Office.Interop.Excel.Worksheet)B).Cells.SparklineGroups)
						{
							A.A(sparklineGroup);
						}
						return;
					}
				}
			}
		}
		IEnumerator enumerator2 = default(IEnumerator);
		try
		{
			enumerator2 = ((Range)B).Cells.SparklineGroups.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				SparklineGroup a2 = (SparklineGroup)enumerator2.Current;
				A.A(a2);
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		finally
		{
			if (enumerator2 is IDisposable)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					(enumerator2 as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	internal static void A(WorksheetItem A, Chart B)
	{
		Application application = B.Application;
		Range range = null;
		IEnumerator enumerator = default(IEnumerator);
		Range range2;
		try
		{
			enumerator = ((IEnumerable)B.FullSeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
			while (enumerator.MoveNext())
			{
				string[] array = ExcelAddIn1.Charts.Helpers.A((Series)enumerator.Current);
				foreach (string cell in array)
				{
					range2 = null;
					try
					{
						range2 = ((_Application)B.Application).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value));
						range2.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						if (range == null)
						{
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
								range = range2;
								break;
							}
						}
						else
						{
							range = application.Union(range, range2, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
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
		if (range != null)
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
			A.F(range);
		}
		range2 = null;
		range = null;
		application = null;
	}

	internal static void B(WorksheetItem A, object B)
	{
		IEnumerator enumerator = default(IEnumerator);
		if (B is Microsoft.Office.Interop.Excel.Worksheet)
		{
			try
			{
				enumerator = ((Microsoft.Office.Interop.Excel.Worksheet)B).Hyperlinks.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Hyperlink b = (Hyperlink)enumerator.Current;
					Other.A(A, b);
				}
				while (true)
				{
					switch (2)
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
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							(enumerator as IDisposable).Dispose();
							goto end_IL_005a;
						}
						continue;
						end_IL_005a:
						break;
					}
				}
			}
		}
		Application application = MH.A.Application;
		IEnumerator enumerator2 = default(IEnumerator);
		try
		{
			enumerator2 = ((Range)B).Worksheet.Hyperlinks.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				Hyperlink hyperlink = (Hyperlink)enumerator2.Current;
				if (application.Intersect(hyperlink.Range, (Range)B, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == null)
				{
					continue;
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
				Other.A(A, hyperlink);
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

	private static void A(WorksheetItem A, Hyperlink B)
	{
		try
		{
			_ = B.TextToDisplay;
			A.A(B);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	internal static void A()
	{
		throw new NotImplementedException();
	}

	internal static void B()
	{
		throw new NotImplementedException();
	}

	internal static void C()
	{
		throw new NotImplementedException();
	}

	internal static void D()
	{
		throw new NotImplementedException();
	}
}
