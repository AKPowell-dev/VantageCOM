using System;
using System.Collections;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

[StandardModule]
internal sealed class IG
{
	public static int A(this Presentations A)
	{
		int num = A.Count;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Presentation presentation = (Presentation)enumerator.Current;
				try
				{
					_ = presentation.Name;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					num = checked(num - 1);
					ProjectData.ClearProjectError();
				}
			}
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
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		return num;
	}
}
