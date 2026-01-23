using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Utilities;

namespace PowerPointAddIn1.Shapes;

public sealed class Stack
{
	internal struct EE
	{
		public Bitmap A;

		public string A;

		public string B;

		public Action A;
	}

	private static EE? m_A;

	internal static EE A
	{
		get
		{
			if (!Stack.m_A.HasValue)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				A();
			}
			return Stack.m_A.Value;
		}
		set
		{
			Stack.m_A = value;
		}
	}

	private static void A(Action A, Bitmap B, string C, string D)
	{
		Stack.A = new EE
		{
			A = A,
			A = new Bitmap(B),
			A = C,
			B = D
		};
		KG.A.InvalidateControl(AH.A(85008));
	}

	private static void A()
	{
		A(F, OB.StackLeft, AH.A(85027), AH.A(85048));
	}

	private static void B()
	{
		A(G, OB.StackRight, AH.A(85452), AH.A(85475));
	}

	private static void C()
	{
		A(H, OB.StackUp, AH.A(85887), AH.A(85904));
	}

	private static void D()
	{
		A(I, OB.StackDown, AH.A(86302), AH.A(86323));
	}

	internal static void E()
	{
		Stack.A.A();
	}

	internal static void F()
	{
		if (!A())
		{
			return;
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
			A();
			ShapeRange shapeRange = Stack.A();
			if (shapeRange == null)
			{
				return;
			}
			while (true)
			{
				switch (6)
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
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					new wpfGapSize(shape, Dialogs.DialogPosition.Above, shapeRange, A).Show();
					_ = null;
					Base.LogActivity(AH.A(85027));
					shape = null;
					shapeRange = null;
					return;
				}
			}
		}
	}

	private static void A(ShapeRange A, float B)
	{
		checked
		{
			try
			{
				List<Shape> ShapesList = Stack.A(A);
				Base.SortShapesByLeftPosition(ref ShapesList);
				List<Shape> list = ShapesList;
				float left = list[0].Left + list[0].Width + B;
				int num = list.Count - 1;
				for (int i = 1; i <= num; i++)
				{
					Shape shape = list[i];
					shape.Left = left;
					left = shape.Left + shape.Width + B;
					shape = null;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					list = null;
					return;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			finally
			{
				A = null;
				List<Shape> ShapesList = null;
			}
		}
	}

	internal static void G()
	{
		if (!A())
		{
			return;
		}
		B();
		ShapeRange shapeRange = Stack.A();
		if (shapeRange == null)
		{
			return;
		}
		NG.A.Application.StartNewUndoEntry();
		Shape shape = shapeRange[1];
		int count = shapeRange.Count;
		for (int i = 2; i <= count; i = checked(i + 1))
		{
			if (shapeRange[i].Left + shapeRange[i].Width > shape.Left + shape.Width)
			{
				shape = shapeRange[i];
			}
		}
		new wpfGapSize(shape, Dialogs.DialogPosition.Above, shapeRange, B).Show();
		_ = null;
		Base.LogActivity(AH.A(85452));
		shape = null;
		shapeRange = null;
	}

	private static void B(ShapeRange A, float B)
	{
		checked
		{
			try
			{
				List<Shape> ShapesList = Stack.A(A);
				Base.SortShapesByRightPosition(ref ShapesList);
				List<Shape> list = ShapesList;
				float num = list[0].Left - B;
				int num2 = list.Count - 1;
				for (int i = 1; i <= num2; i++)
				{
					Shape shape = list[i];
					shape.Left = num - shape.Width;
					num = shape.Left - B;
					shape = null;
				}
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
					list = null;
					return;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			finally
			{
				A = null;
				List<Shape> ShapesList = null;
			}
		}
	}

	internal static void H()
	{
		if (!A())
		{
			return;
		}
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
			C();
			ShapeRange shapeRange = Stack.A();
			if (shapeRange == null)
			{
				return;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
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
						switch (5)
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
					switch (4)
					{
					case 0:
						continue;
					}
					new wpfGapSize(shape, Dialogs.DialogPosition.Above, shapeRange, C).Show();
					_ = null;
					Base.LogActivity(AH.A(85887));
					shape = null;
					shapeRange = null;
					return;
				}
			}
		}
	}

	private static void C(ShapeRange A, float B)
	{
		checked
		{
			try
			{
				List<Shape> ShapesList = Stack.A(A);
				Base.SortShapesByTopPosition(ref ShapesList);
				List<Shape> list = ShapesList;
				float top = list[0].Top + list[0].Height + B;
				int num = list.Count - 1;
				for (int i = 1; i <= num; i++)
				{
					Shape shape = list[i];
					shape.Top = top;
					top = shape.Top + shape.Height + B;
					shape = null;
				}
				list = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			finally
			{
				A = null;
				List<Shape> ShapesList = null;
			}
		}
	}

	internal static void I()
	{
		if (!A())
		{
			return;
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			D();
			ShapeRange shapeRange = Stack.A();
			if (shapeRange == null)
			{
				return;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				NG.A.Application.StartNewUndoEntry();
				Shape shape = shapeRange[1];
				int count = shapeRange.Count;
				for (int i = 2; i <= count; i = checked(i + 1))
				{
					if (shapeRange[i].Top + shapeRange[i].Height > shape.Top + shape.Height)
					{
						shape = shapeRange[i];
					}
				}
				new wpfGapSize(shape, Dialogs.DialogPosition.Below, shapeRange, D).Show();
				_ = null;
				Base.LogActivity(AH.A(86302));
				shape = null;
				shapeRange = null;
				return;
			}
		}
	}

	private static void D(ShapeRange A, float B)
	{
		checked
		{
			try
			{
				List<Shape> ShapesList = Stack.A(A);
				Base.SortShapesByBottomPosition(ref ShapesList);
				List<Shape> list = ShapesList;
				float num = list[0].Top - B;
				int num2 = list.Count - 1;
				for (int i = 1; i <= num2; i++)
				{
					Shape shape = list[i];
					shape.Top = num - shape.Height;
					num = shape.Top - B;
					shape = null;
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					list = null;
					return;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			finally
			{
				A = null;
				List<Shape> ShapesList = null;
			}
		}
	}

	private static bool A()
	{
		return Licensing.AllowRestrictedMode();
	}

	internal static ShapeRange A()
	{
		try
		{
			ShapeRange shapeRange = Base.SelectedShapes();
			if (shapeRange.Count > 1)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return shapeRange;
				}
			}
			Helpers.TwoOrMoreShapesRequiredError();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Helpers.TwoOrMoreShapesRequiredError();
			ProjectData.ClearProjectError();
		}
		finally
		{
			ShapeRange shapeRange = null;
		}
		return null;
	}

	private static List<Shape> A(ShapeRange A)
	{
		List<Shape> list = new List<Shape>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Shape item = (Shape)enumerator.Current;
				list.Add(item);
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
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		return list;
	}
}
