using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

[StandardModule]
internal sealed class MC
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
			A = null;
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

	public static CustomTaskPane A(string A)
	{
		CustomTaskPaneCollection customTaskPanes = PC.A.CustomTaskPanes;
		checked
		{
			for (int i = customTaskPanes.Count - 1; i >= 0; i += -1)
			{
				if (Operators.CompareString(customTaskPanes[i].Title, A, TextCompare: false) != 0)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return customTaskPanes[i];
				}
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				customTaskPanes = null;
				return null;
			}
		}
	}

	public static void A(string A)
	{
		checked
		{
			try
			{
				CustomTaskPaneCollection customTaskPanes = PC.A.CustomTaskPanes;
				for (int i = customTaskPanes.Count - 1; i >= 0; i += -1)
				{
					try
					{
						if (Operators.CompareString(customTaskPanes[i].Title, A, TextCompare: false) == 0)
						{
							customTaskPanes.RemoveAt(i);
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				customTaskPanes = null;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
	}
}
