using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

[StandardModule]
internal sealed class JG
{
	public static void A(object A)
	{
		try
		{
			int num = 0;
			do
			{
				num = Marshal.ReleaseComObject(RuntimeHelpers.GetObjectValue(A));
			}
			while (num > 0);
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
				A = null;
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			A = null;
			ProjectData.ClearProjectError();
		}
		finally
		{
			GC.Collect();
			GC.WaitForPendingFinalizers();
			GC.Collect();
			GC.WaitForPendingFinalizers();
		}
	}
}
