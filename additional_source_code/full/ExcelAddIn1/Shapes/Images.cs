using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Shapes;

public sealed class Images
{
	internal static readonly int A = 28;

	internal static readonly int B = 29;

	internal static bool A(Shape A)
	{
		bool result = false;
		try
		{
			MsoShapeType type = A.Type;
			if (type == MsoShapeType.msoLinkedPicture)
			{
				goto IL_0028;
			}
			if (type == MsoShapeType.msoPicture)
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
				goto IL_0028;
			}
			if (A.Type == (MsoShapeType)Images.A)
			{
				goto IL_005a;
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
			if (A.Type == (MsoShapeType)B)
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
				goto IL_005a;
			}
			goto end_IL_0002;
			IL_005a:
			result = true;
			goto end_IL_0002;
			IL_0028:
			result = true;
			end_IL_0002:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}
}
