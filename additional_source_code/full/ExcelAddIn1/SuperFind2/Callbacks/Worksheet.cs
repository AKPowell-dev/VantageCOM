using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.SuperFind2.Results;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class Worksheet
{
	internal static void A(WorksheetItem A, Microsoft.Office.Interop.Excel.Worksheet B)
	{
		try
		{
			int num = Conversions.ToInteger(NewLateBinding.LateGet(B.ChartObjects(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(52690), new object[0], null, null, null));
			for (int i = 1; i <= num; i = checked(i + 1))
			{
				A.A((ChartObject)B.ChartObjects(i));
			}
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
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	internal static void B(WorksheetItem A, Microsoft.Office.Interop.Excel.Worksheet B)
	{
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = B.Shapes.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Shape shape = (Shape)enumerator.Current;
					if (shape.HasChart == MsoTriState.msoTrue)
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
					if (shape.Type == MsoShapeType.msoComment)
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
						break;
					}
					A.A(shape);
				}
				while (true)
				{
					switch (2)
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
