using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Library2;

public sealed class Slides
{
	[CompilerGenerated]
	internal sealed class ID
	{
		public Microsoft.Office.Interop.PowerPoint.Presentation A;

		public ID(ID A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal void A()
		{
			this.A.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).Copy();
		}
	}

	internal static int A(Microsoft.Office.Interop.PowerPoint.Presentation A, Microsoft.Office.Interop.PowerPoint.Application B)
	{
		ID a = default(ID);
		ID CS_0024_003C_003E8__locals3 = new ID(a);
		CS_0024_003C_003E8__locals3.A = A;
		int count = CS_0024_003C_003E8__locals3.A.Slides.Count;
		clsClipboard.CopyWithWait((Action)([SpecialName] () =>
		{
			CS_0024_003C_003E8__locals3.A.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).Copy();
		}), 4000);
		int num = 1;
		while (true)
		{
			try
			{
				B.CommandBars.ExecuteMso(AH.A(58900));
				System.Windows.Forms.Application.DoEvents();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Thread.Sleep(100);
				if (num == 10)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						throw;
					}
				}
				ProjectData.ClearProjectError();
				goto IL_008e;
			}
			break;
			IL_008e:
			num = checked(num + 1);
			if (num <= 10)
			{
				continue;
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
			break;
		}
		clsClipboard.ClearClipboard();
		return count;
	}
}
