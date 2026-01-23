using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Shapes;

public sealed class Images
{
	internal static readonly int A = 28;

	internal static readonly int B = 29;

	internal static readonly int C = 17;

	internal static readonly int D = 18;

	internal static bool A(InlineShape A)
	{
		bool result = false;
		try
		{
			WdInlineShapeType type = A.Type;
			if ((uint)(type - 3) <= 1u)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					result = true;
					break;
				}
			}
			else
			{
				if (A.Type == (WdInlineShapeType)C)
				{
					goto IL_005a;
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
				if (A.Type == (WdInlineShapeType)D)
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
					goto IL_005a;
				}
			}
			goto end_IL_0002;
			IL_005a:
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

	internal static bool A(Microsoft.Office.Interop.Word.Shape A)
	{
		bool result = false;
		try
		{
			MsoShapeType type = A.Type;
			if (type == MsoShapeType.msoLinkedPicture || type == MsoShapeType.msoPicture)
			{
				result = true;
			}
			else
			{
				if (A.Type == (MsoShapeType)Images.A)
				{
					goto IL_0052;
				}
				while (true)
				{
					switch (5)
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
				if (A.Type == (MsoShapeType)Images.B)
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
					goto IL_0052;
				}
			}
			goto end_IL_0002;
			IL_0052:
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

	internal static bool B(Microsoft.Office.Interop.Word.Shape A)
	{
		bool result = false;
		try
		{
			if (A.Type == (MsoShapeType)Images.A)
			{
				goto IL_0031;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (A.Type == (MsoShapeType)Images.B)
			{
				goto IL_0031;
			}
			goto end_IL_0002;
			IL_0031:
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

	internal static bool B(InlineShape A)
	{
		bool result = false;
		try
		{
			if (A.Type == (WdInlineShapeType)C)
			{
				goto IL_0033;
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
			if (A.Type == (WdInlineShapeType)D)
			{
				goto IL_0033;
			}
			goto end_IL_0002;
			IL_0033:
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
