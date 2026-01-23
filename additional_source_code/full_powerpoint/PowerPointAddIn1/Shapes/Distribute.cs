using System;
using System.Collections;
using System.Collections.Generic;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Utilities;

namespace PowerPointAddIn1.Shapes;

public sealed class Distribute
{
	internal static void A()
	{
		if (!A())
		{
			return;
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
			ShapeRange shapeRange = Stack.A();
			if (shapeRange == null)
			{
				return;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				NG.A.Application.StartNewUndoEntry();
				Shape shape = shapeRange[1];
				int count = shapeRange.Count;
				for (int i = 2; i <= count; i = checked(i + 1))
				{
					if (shapeRange[i].Left < shape.Left)
					{
						shape = shapeRange[i];
					}
				}
				new wpfGapSize(shape, Dialogs.DialogPosition.Above, shapeRange, A).Show();
				_ = null;
				Base.LogActivity(AH.A(81880));
				shape = null;
				shapeRange = null;
				return;
			}
		}
	}

	private static void A(ShapeRange A, float B)
	{
		int count = A.Count;
		float num = 10000f;
		float num2 = -10000f;
		List<Shape> ShapesList = new List<Shape>();
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Shape shape = (Shape)enumerator.Current;
					Shape shape2 = shape;
					if (shape2.Left < num)
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
						num = shape2.Left;
					}
					if (shape2.Left + shape2.Width > num2)
					{
						num2 = shape2.Left + shape2.Width;
					}
					shape2 = null;
					ShapesList.Add(shape);
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_00a7;
					}
					continue;
					end_IL_00a7:
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (4)
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
			ShapesList = null;
			ProjectData.ClearProjectError();
			return;
		}
		float num3 = (num2 - num - B * (float)checked(count - 1)) / (float)count;
		Base.SortShapesByLeftPosition(ref ShapesList);
		using (List<Shape>.Enumerator enumerator2 = ShapesList.GetEnumerator())
		{
			while (enumerator2.MoveNext())
			{
				Shape current = enumerator2.Current;
				current.Left = num;
				current.Width = num3;
				num = num + num3 + B;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_012c;
				}
				continue;
				end_IL_012c:
				break;
			}
		}
		ShapesList = null;
	}

	internal static void B()
	{
		if (!A())
		{
			return;
		}
		ShapeRange shapeRange = Stack.A();
		if (shapeRange == null)
		{
			return;
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
			NG.A.Application.StartNewUndoEntry();
			Shape shape = shapeRange[1];
			int count = shapeRange.Count;
			for (int i = 2; i <= count; i = checked(i + 1))
			{
				if (!(shapeRange[i].Top < shape.Top))
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
					break;
				}
				shape = shapeRange[i];
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				new wpfGapSize(shape, Dialogs.DialogPosition.Above, shapeRange, B).Show();
				_ = null;
				Base.LogActivity(AH.A(81927));
				shape = null;
				shapeRange = null;
				return;
			}
		}
	}

	private static void B(ShapeRange A, float B)
	{
		int count = A.Count;
		float num = 10000f;
		float num2 = -10000f;
		List<Shape> ShapesList = new List<Shape>();
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Shape shape = (Shape)enumerator.Current;
					Shape shape2 = shape;
					if (shape2.Top < num)
					{
						num = shape2.Top;
					}
					if (shape2.Top + shape2.Height > num2)
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
						num2 = shape2.Top + shape2.Height;
					}
					shape2 = null;
					ShapesList.Add(shape);
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_009e;
					}
					continue;
					end_IL_009e:
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (1)
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
			ShapesList = null;
			ProjectData.ClearProjectError();
			return;
		}
		float num3 = (num2 - num - B * (float)checked(count - 1)) / (float)count;
		Base.SortShapesByTopPosition(ref ShapesList);
		foreach (Shape item in ShapesList)
		{
			item.Top = num;
			item.Height = num3;
			num = num + num3 + B;
		}
		ShapesList = null;
	}

	private static bool A()
	{
		return Licensing.AllowRestrictedMode();
	}
}
