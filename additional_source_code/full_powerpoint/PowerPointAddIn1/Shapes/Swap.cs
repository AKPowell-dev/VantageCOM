using System;
using System.Drawing;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes;

public sealed class Swap
{
	internal struct FE
	{
		public Bitmap A;

		public string A;

		public string B;

		public Action A;
	}

	private static FE? m_A;

	internal static FE A
	{
		get
		{
			if (!Swap.m_A.HasValue)
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
				A();
			}
			return Swap.m_A.Value;
		}
		set
		{
			Swap.m_A = value;
		}
	}

	private static void A(Action A, Bitmap B, string C, string D)
	{
		Swap.A = new FE
		{
			A = A,
			A = new Bitmap(B),
			A = C,
			B = D
		};
		KG.A.InvalidateControl(AH.A(86725));
	}

	private static void A()
	{
		A(G, OB.AnchorTopLeft, AH.A(86742), AH.A(86783));
	}

	private static void B()
	{
		A(H, OB.AnchorTopRight, AH.A(86884), AH.A(86927));
	}

	private static void C()
	{
		A(I, OB.AnchorBottomLeft, AH.A(87030), AH.A(87077));
	}

	private static void D()
	{
		A(J, OB.AnchorBottomRight, AH.A(87184), AH.A(87233));
	}

	private static void E()
	{
		A(K, OB.AnchorCenter, AH.A(87342), AH.A(87379));
	}

	internal static void F()
	{
		Swap.A.A();
	}

	internal static void G()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		A();
		try
		{
			ShapeRange shapeRange = Base.SelectedShapes();
			if (shapeRange.Count == 2)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				NG.A.Application.StartNewUndoEntry();
				float top = shapeRange[1].Top;
				float left = shapeRange[1].Left;
				shapeRange[1].Top = shapeRange[2].Top;
				shapeRange[1].Left = shapeRange[2].Left;
				shapeRange[2].Top = top;
				shapeRange[2].Left = left;
			}
			else
			{
				L();
			}
			shapeRange = null;
			Base.LogActivity(AH.A(86742));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			L();
			ProjectData.ClearProjectError();
		}
	}

	internal static void H()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		B();
		try
		{
			ShapeRange shapeRange = Base.SelectedShapes();
			if (shapeRange.Count == 2)
			{
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
				NG.A.Application.StartNewUndoEntry();
				float top = shapeRange[1].Top;
				float num = shapeRange[1].Left + shapeRange[1].Width;
				shapeRange[1].Top = shapeRange[2].Top;
				shapeRange[1].Left = shapeRange[2].Left + shapeRange[2].Width - shapeRange[1].Width;
				shapeRange[2].Top = top;
				shapeRange[2].Left = num - shapeRange[2].Width;
			}
			else
			{
				L();
			}
			shapeRange = null;
			Base.LogActivity(AH.A(86884));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			L();
			ProjectData.ClearProjectError();
		}
	}

	internal static void I()
	{
		if (!Licensing.AllowRestrictedMode())
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
			try
			{
				ShapeRange shapeRange = Base.SelectedShapes();
				if (shapeRange.Count == 2)
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
					NG.A.Application.StartNewUndoEntry();
					float num = shapeRange[1].Top + shapeRange[1].Height;
					float left = shapeRange[1].Left;
					shapeRange[1].Top = shapeRange[2].Top + shapeRange[2].Height - shapeRange[1].Height;
					shapeRange[1].Left = shapeRange[2].Left;
					shapeRange[2].Top = num - shapeRange[2].Height;
					shapeRange[2].Left = left;
				}
				else
				{
					L();
				}
				shapeRange = null;
				Base.LogActivity(AH.A(87030));
				return;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				L();
				ProjectData.ClearProjectError();
				return;
			}
		}
	}

	internal static void J()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		D();
		try
		{
			ShapeRange shapeRange = Base.SelectedShapes();
			if (shapeRange.Count == 2)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				NG.A.Application.StartNewUndoEntry();
				float num = shapeRange[1].Top + shapeRange[1].Height;
				float num2 = shapeRange[1].Left + shapeRange[1].Width;
				shapeRange[1].Top = shapeRange[2].Top + shapeRange[2].Height - shapeRange[1].Height;
				shapeRange[1].Left = shapeRange[2].Left + shapeRange[2].Width - shapeRange[1].Width;
				shapeRange[2].Top = num - shapeRange[2].Height;
				shapeRange[2].Left = num2 - shapeRange[2].Width;
			}
			else
			{
				L();
			}
			shapeRange = null;
			Base.LogActivity(AH.A(87184));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			L();
			ProjectData.ClearProjectError();
		}
	}

	internal static void K()
	{
		if (!Licensing.AllowRestrictedMode())
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
			E();
			try
			{
				ShapeRange shapeRange = Base.SelectedShapes();
				if (shapeRange.Count == 2)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
					NG.A.Application.StartNewUndoEntry();
					float num = shapeRange[1].Top + shapeRange[1].Height / 2f;
					float num2 = shapeRange[1].Left + shapeRange[1].Width / 2f;
					shapeRange[1].Top = shapeRange[2].Top + shapeRange[2].Height / 2f - shapeRange[1].Height / 2f;
					shapeRange[1].Left = shapeRange[2].Left + shapeRange[2].Width / 2f - shapeRange[1].Width / 2f;
					shapeRange[2].Top = num - shapeRange[2].Height / 2f;
					shapeRange[2].Left = num2 - shapeRange[2].Width / 2f;
				}
				else
				{
					L();
				}
				shapeRange = null;
				Base.LogActivity(AH.A(87342));
				return;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				L();
				ProjectData.ClearProjectError();
				return;
			}
		}
	}

	private static void L()
	{
		Forms.WarningMessage(AH.A(87476));
	}
}
