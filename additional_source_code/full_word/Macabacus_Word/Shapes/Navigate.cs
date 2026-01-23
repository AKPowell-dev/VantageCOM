using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Shapes;

public sealed class Navigate
{
	internal static void A(object A)
	{
		if (A is ContentControl)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					ContentControl obj = (ContentControl)A;
					Navigate.A(obj.Range);
					obj.Range.Select();
					_ = null;
					return;
				}
				}
			}
		}
		if (A is InlineShape)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					Navigate.A(((InlineShape)A).Range);
					NewLateBinding.LateCall(A, null, XC.A(12515), new object[0], null, null, null, IgnoreReturn: true);
					_ = null;
					return;
				}
			}
		}
		if (A is Table)
		{
			Navigate.A(((Table)A).Range);
			NewLateBinding.LateCall(A, null, XC.A(12515), new object[0], null, null, null, IgnoreReturn: true);
			_ = null;
		}
		else if (A is Shape)
		{
			Navigate.A(((Shape)A).Anchor);
			NewLateBinding.LateCall(A, null, XC.A(12515), new object[0], null, null, null, IgnoreReturn: true);
			_ = null;
		}
	}

	private static void A(Range A)
	{
		if (!IsInHeaderFooter(A))
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Window activeWindow = A.Application.ActiveWindow;
					object Start = RuntimeHelpers.GetObjectValue(Missing.Value);
					activeWindow.ScrollIntoView(A, ref Start);
					return;
				}
				}
			}
		}
		B(A);
	}

	public static bool IsInHeaderFooter(Shape shp)
	{
		return IsInHeaderFooter(shp.Anchor);
	}

	public static bool IsInHeaderFooter(InlineShape shp)
	{
		return IsInHeaderFooter(shp.Range);
	}

	public static bool IsInHeaderFooter(Table tbl)
	{
		return IsInHeaderFooter(tbl.Range);
	}

	public static bool IsInHeaderFooter(ContentControl cc)
	{
		return IsInHeaderFooter(cc.Range);
	}

	public static bool IsInHeaderFooter(Range rng)
	{
		WdStoryType storyType = rng.StoryType;
		if ((uint)(storyType - 6) <= 5u)
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
					return true;
				}
			}
		}
		return false;
	}

	internal static void B(Range A)
	{
		View view = A.Application.ActiveWindow.View;
		if (view.Type != WdViewType.wdNormalView)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			object left = A.get_Information(WdInformation.wdHeaderFooterType);
			if (Conversions.ToBoolean(Conversions.ToBoolean(Operators.CompareObjectEqual(left, 0, TextCompare: false)) || Conversions.ToBoolean(Operators.CompareObjectEqual(left, 1, TextCompare: false)) || Conversions.ToBoolean(Operators.CompareObjectEqual(left, 4, TextCompare: false))))
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
				view.SeekView = WdSeekView.wdSeekCurrentPageHeader;
			}
			else
			{
				int num;
				if (!Conversions.ToBoolean(Operators.CompareObjectEqual(left, 2, TextCompare: false)) && !Conversions.ToBoolean(Operators.CompareObjectEqual(left, 3, TextCompare: false)))
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
					num = (Conversions.ToBoolean(Operators.CompareObjectEqual(left, 5, TextCompare: false)) ? 1 : 0);
				}
				else
				{
					num = 1;
				}
				if (Conversions.ToBoolean((byte)num != 0))
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
					view.SeekView = WdSeekView.wdSeekCurrentPageFooter;
				}
			}
		}
		view = null;
	}

	internal static bool A(Application A)
	{
		try
		{
			WdSeekView seekView = A.ActiveWindow.View.SeekView;
			if ((uint)(seekView - 1) <= 5u)
			{
				goto IL_003d;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if ((uint)(seekView - 9) <= 1u)
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
				goto IL_003d;
			}
			goto end_IL_0000;
			IL_003d:
			return true;
			end_IL_0000:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return false;
	}
}
