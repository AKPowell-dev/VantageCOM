using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

internal sealed class QB
{
	internal static List<object> A(Worksheet A)
	{
		List<object> list = new List<object>();
		foreach (object item in (IEnumerable)NewLateBinding.LateGet(A, null, VH.A(8668), new object[0], null, null, null))
		{
			object objectValue = RuntimeHelpers.GetObjectValue(item);
			list.Add(RuntimeHelpers.GetObjectValue(objectValue));
		}
		return list;
	}

	internal static Range A(object A)
	{
		return NewLateBinding.LateGet(A, null, VH.A(8701), new object[0], null, null, null) as Range;
	}

	internal static bool A(Range A)
	{
		return object.Equals(RuntimeHelpers.GetObjectValue(A.HasArray), true);
	}

	internal static string A(Range A)
	{
		return NewLateBinding.LateGet(A, null, VH.A(1998), new object[0], null, null, null).ToString();
	}

	internal static string B(Range A)
	{
		return NewLateBinding.LateGet(A, null, VH.A(8714), new object[0], null, null, null).ToString();
	}

	internal static Range A(Range A, long B)
	{
		return (Range)A.Cells.get_Item((object)B, RuntimeHelpers.GetObjectValue(Missing.Value));
	}

	internal static Worksheet A(Workbook A, string B)
	{
		return (Worksheet)A.Worksheets.get_Item((object)B);
	}
}
