using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.Shapes;

public sealed class Base
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Comparison<Shape> A;

		public static Comparison<Shape> B;

		public static Comparison<Shape> C;

		public static Comparison<Shape> D;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal int A(Shape A, Shape B)
		{
			if (A.Left < B.Left)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return -1;
					}
				}
			}
			return 1;
		}

		[SpecialName]
		internal int B(Shape A, Shape B)
		{
			if (A.Left + A.Width > B.Left + B.Width)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return -1;
					}
				}
			}
			return 1;
		}

		[SpecialName]
		internal int C(Shape A, Shape B)
		{
			if (A.Top < B.Top)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return -1;
					}
				}
			}
			return 1;
		}

		[SpecialName]
		internal int D(Shape A, Shape B)
		{
			if (A.Top + A.Height > B.Top + B.Height)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return -1;
					}
				}
			}
			return 1;
		}
	}

	public static ShapeRange SelectedShapes()
	{
		return SelectedShapes(NG.A.Application.ActiveWindow.Selection);
	}

	public static ShapeRange SelectedShapes(Selection sel)
	{
		if (sel.HasChildShapeRange)
		{
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
					return sel.ChildShapeRange;
				}
			}
		}
		return sel.ShapeRange;
	}

	public static ShapeRange SelectedShapes2()
	{
		return NG.A.Application.ActiveWindow.Selection.ShapeRange;
	}

	public static void SortShapesByLeftPosition(ref List<Shape> ShapesList)
	{
		ShapesList.Sort([SpecialName] (Shape A, Shape B) =>
		{
			if (A.Left < B.Left)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return -1;
					}
				}
			}
			return 1;
		});
	}

	public static void SortShapesByRightPosition(ref List<Shape> ShapesList)
	{
		List<Shape> obj = ShapesList;
		Comparison<Shape> comparison;
		if (_Closure_0024__.B == null)
		{
			comparison = (_Closure_0024__.B = [SpecialName] (Shape A, Shape B) =>
			{
				if (A.Left + A.Width > B.Left + B.Width)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							return -1;
						}
					}
				}
				return 1;
			});
		}
		else
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			comparison = _Closure_0024__.B;
		}
		obj.Sort(comparison);
	}

	public static void SortShapesByTopPosition(ref List<Shape> ShapesList)
	{
		List<Shape> obj = ShapesList;
		Comparison<Shape> comparison;
		if (_Closure_0024__.C == null)
		{
			comparison = (_Closure_0024__.C = [SpecialName] (Shape A, Shape B) =>
			{
				if (A.Top < B.Top)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							return -1;
						}
					}
				}
				return 1;
			});
		}
		else
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
			comparison = _Closure_0024__.C;
		}
		obj.Sort(comparison);
	}

	public static void SortShapesByBottomPosition(ref List<Shape> ShapesList)
	{
		List<Shape> obj = ShapesList;
		Comparison<Shape> comparison;
		if (_Closure_0024__.D == null)
		{
			comparison = (_Closure_0024__.D = [SpecialName] (Shape A, Shape B) =>
			{
				if (A.Top + A.Height > B.Top + B.Height)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							return -1;
						}
					}
				}
				return 1;
			});
		}
		else
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			comparison = _Closure_0024__.D;
		}
		obj.Sort(comparison);
	}

	public static void AlignError()
	{
		Forms.WarningMessage(AH.A(81244));
	}

	public static void LogActivity(string strActivity)
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)1, strActivity);
	}
}
