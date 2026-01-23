using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1.Shapes;

public sealed class Align
{
	internal struct ZD
	{
		public string A;

		public string B;

		public string C;

		public Action A;
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Microsoft.Office.Interop.PowerPoint.Shape, double> A;

		public static Func<Microsoft.Office.Interop.PowerPoint.Shape, Microsoft.Office.Interop.PowerPoint.Shape> A;

		public static Func<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>, BB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>> A;

		public static Func<BB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>, int> A;

		public static Func<BB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>, CB<double, int>> A;

		public static Func<Microsoft.Office.Interop.PowerPoint.Shape, double> B;

		public static Func<Microsoft.Office.Interop.PowerPoint.Shape, Microsoft.Office.Interop.PowerPoint.Shape> B;

		public static Func<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>, DB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>> A;

		public static Func<DB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>, int> A;

		public static Func<DB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>, EB<double, int>> A;

		public static Func<Microsoft.Office.Interop.PowerPoint.Shape, double> C;

		public static Func<Microsoft.Office.Interop.PowerPoint.Shape, Microsoft.Office.Interop.PowerPoint.Shape> C;

		public static Func<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>, FB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>> A;

		public static Func<FB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>, int> A;

		public static Func<FB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>, GB<double, int>> A;

		public static Func<Microsoft.Office.Interop.PowerPoint.Shape, double> D;

		public static Func<Microsoft.Office.Interop.PowerPoint.Shape, Microsoft.Office.Interop.PowerPoint.Shape> D;

		public static Func<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>, HB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>> A;

		public static Func<HB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>, int> A;

		public static Func<HB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>, IB<double, int>> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal double A(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return Math.Round(A.Top, MISALIGNMENT_ROUND);
		}

		[SpecialName]
		internal Microsoft.Office.Interop.PowerPoint.Shape A(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return A;
		}

		[SpecialName]
		internal BB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> A(double A, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape> B)
		{
			return new BB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>(A, B);
		}

		[SpecialName]
		internal int A(BB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> A)
		{
			return A.Group.Count();
		}

		[SpecialName]
		internal CB<double, int> A(BB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> A)
		{
			return new CB<double, int>(A.top, A.Group.Count());
		}

		[SpecialName]
		internal double B(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return Math.Round(A.Left, MISALIGNMENT_ROUND);
		}

		[SpecialName]
		internal Microsoft.Office.Interop.PowerPoint.Shape B(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return A;
		}

		[SpecialName]
		internal DB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> A(double A, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape> B)
		{
			return new DB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>(A, B);
		}

		[SpecialName]
		internal int A(DB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> A)
		{
			return A.Group.Count();
		}

		[SpecialName]
		internal EB<double, int> A(DB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> A)
		{
			return new EB<double, int>(A.left, A.Group.Count());
		}

		[SpecialName]
		internal double C(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return Math.Round(A.Top + A.Height, MISALIGNMENT_ROUND);
		}

		[SpecialName]
		internal Microsoft.Office.Interop.PowerPoint.Shape C(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return A;
		}

		[SpecialName]
		internal FB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> A(double A, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape> B)
		{
			return new FB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>(A, B);
		}

		[SpecialName]
		internal int A(FB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> A)
		{
			return A.Group.Count();
		}

		[SpecialName]
		internal GB<double, int> A(FB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> A)
		{
			return new GB<double, int>(A.bottom, A.Group.Count());
		}

		[SpecialName]
		internal double D(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return Math.Round(A.Left + A.Width, MISALIGNMENT_ROUND);
		}

		[SpecialName]
		internal Microsoft.Office.Interop.PowerPoint.Shape D(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return A;
		}

		[SpecialName]
		internal HB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> A(double A, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape> B)
		{
			return new HB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>(A, B);
		}

		[SpecialName]
		internal int A(HB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> A)
		{
			return A.Group.Count();
		}

		[SpecialName]
		internal IB<double, int> A(HB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> A)
		{
			return new IB<double, int>(A.right, A.Group.Count());
		}
	}

	[CompilerGenerated]
	internal sealed class AE
	{
		public double A;

		public int A;

		[SpecialName]
		internal bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return Align.A(this.A, Math.Round(A.Top, MISALIGNMENT_ROUND), this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class BE
	{
		public double A;

		public int A;

		[SpecialName]
		internal bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return Align.A(this.A, Math.Round(A.Left, MISALIGNMENT_ROUND), this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class CE
	{
		public double A;

		public int A;

		[SpecialName]
		internal bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return Align.A(this.A, Math.Round(A.Top + A.Height, MISALIGNMENT_ROUND), this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class DE
	{
		public double A;

		public int A;

		[SpecialName]
		internal bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return Align.A(this.A, Math.Round(A.Left + A.Width, MISALIGNMENT_ROUND), this.A);
		}
	}

	private static ZD? m_A;

	public static readonly int MISALIGNMENT_ROUND = 2;

	internal static ZD A
	{
		get
		{
			if (!Align.m_A.HasValue)
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
				A();
			}
			return Align.m_A.Value;
		}
		set
		{
			Align.m_A = value;
		}
	}

	private static void A(Action A, string B, string C, string D)
	{
		Align.A = new ZD
		{
			A = A,
			A = B,
			B = C,
			C = D
		};
		KG.A.InvalidateControl(AH.A(75850));
	}

	private static void A()
	{
		A(H, AH.A(75869), AH.A(75902), AH.A(75923));
	}

	private static void B()
	{
		A(I, AH.A(76805), AH.A(76840), AH.A(76863));
	}

	private static void C()
	{
		A(J, AH.A(77759), AH.A(77790), AH.A(77809));
	}

	private static void D()
	{
		A(K, AH.A(78677), AH.A(78714), AH.A(78739));
	}

	private static void E()
	{
		A(L, AH.A(79649), AH.A(79706), AH.A(79745));
	}

	private static void F()
	{
		A(M, AH.A(80099), AH.A(80152), AH.A(80187));
	}

	internal static void G()
	{
		Align.A.A();
	}

	internal static void H()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			A();
			Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
			try
			{
				Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = Base.SelectedShapes();
				application.StartNewUndoEntry();
				RectangleF? rectangleF;
				List<float> list;
				if (shapeRange.Count > 1)
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
					list = new List<float>();
					{
						enumerator = Base.SelectedShapes().GetEnumerator();
						try
						{
							while (enumerator.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
								list.Add(shape.Left);
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_008e;
								}
								continue;
								end_IL_008e:
								break;
							}
						}
						finally
						{
							IDisposable disposable = enumerator as IDisposable;
							if (disposable != null)
							{
								disposable.Dispose();
							}
						}
					}
					list = list.Distinct().ToList();
					if (list.Count > 1)
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
						shapeRange.Left = shapeRange[1].Left;
						goto IL_0169;
					}
					rectangleF = A(application.ActivePresentation);
					if (rectangleF.HasValue)
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
						if (list[0] != rectangleF.Value.Left)
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
							shapeRange.Left = rectangleF.Value.Left;
							goto IL_0161;
						}
					}
					shapeRange.Left = 0f;
					goto IL_0161;
				}
				RectangleF? rectangleF2 = A(application.ActivePresentation);
				if (rectangleF2.HasValue)
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
					if (shapeRange.Left != rectangleF2.Value.Left)
					{
						shapeRange.Left = rectangleF2.Value.Left;
						goto IL_01d4;
					}
				}
				shapeRange.Left = 0f;
				goto IL_01d4;
				IL_01d4:
				rectangleF2 = null;
				goto IL_01dc;
				IL_0161:
				rectangleF = null;
				goto IL_0169;
				IL_0169:
				list = null;
				goto IL_01dc;
				IL_01dc:
				shapeRange = null;
				Base.LogActivity(AH.A(75902));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Base.AlignError();
				ProjectData.ClearProjectError();
			}
			application = null;
			return;
		}
	}

	internal static void I()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
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
			B();
			Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
			try
			{
				Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = Base.SelectedShapes();
				application.StartNewUndoEntry();
				float num;
				List<float> list;
				RectangleF? rectangleF;
				Microsoft.Office.Interop.PowerPoint.Presentation activePresentation;
				if (shapeRange.Count > 1)
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
					list = new List<float>();
					try
					{
						enumerator = shapeRange.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
							list.Add((float)Math.Round(shape.Left + shape.Width, 4));
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_00a1;
							}
							continue;
							end_IL_00a1:
							break;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								(enumerator as IDisposable).Dispose();
								break;
							}
						}
					}
					list = list.Distinct().ToList();
					if (list.Count > 1)
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
						num = shapeRange[1].Left + shapeRange[1].Width;
						goto IL_01ca;
					}
					activePresentation = application.ActivePresentation;
					rectangleF = A(activePresentation);
					if (rectangleF.HasValue)
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
						if ((double)list[0] != Math.Round(rectangleF.Value.Left + rectangleF.Value.Width, 4))
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
							num = rectangleF.Value.Left + rectangleF.Value.Width;
							goto IL_01be;
						}
					}
					num = activePresentation.PageSetup.SlideWidth;
					goto IL_01be;
				}
				Microsoft.Office.Interop.PowerPoint.Presentation activePresentation2 = application.ActivePresentation;
				RectangleF? rectangleF2 = A(activePresentation2);
				Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange2 = shapeRange;
				if (rectangleF2.HasValue && Math.Round(shapeRange2.Left, 4) != Math.Round(rectangleF2.Value.Left + rectangleF2.Value.Width - shapeRange2.Width, 4))
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
					shapeRange2.Left = rectangleF2.Value.Left + rectangleF2.Value.Width - shapeRange2.Width;
				}
				else
				{
					shapeRange2.Left = activePresentation2.PageSetup.SlideWidth - shapeRange.Width;
				}
				shapeRange2 = null;
				rectangleF2 = null;
				activePresentation2 = null;
				goto IL_0316;
				IL_0316:
				Base.LogActivity(AH.A(76840));
				goto end_IL_0033;
				IL_01ca:
				try
				{
					enumerator2 = shapeRange.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
						shape2.Left = num - shape2.Width;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_0203;
						}
						continue;
						end_IL_0203:
						break;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
				list = null;
				goto IL_0316;
				IL_01be:
				rectangleF = null;
				activePresentation = null;
				goto IL_01ca;
				end_IL_0033:;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Base.AlignError();
				ProjectData.ClearProjectError();
			}
			application = null;
			return;
		}
	}

	internal static void J()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
			try
			{
				Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = Base.SelectedShapes();
				application.StartNewUndoEntry();
				RectangleF? rectangleF;
				List<float> list;
				if (shapeRange.Count > 1)
				{
					list = new List<float>();
					try
					{
						enumerator = Base.SelectedShapes().GetEnumerator();
						while (enumerator.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
							list.Add(shape.Top);
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_0084;
							}
							continue;
							end_IL_0084:
							break;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								(enumerator as IDisposable).Dispose();
								break;
							}
						}
					}
					list = list.Distinct().ToList();
					if (list.Count > 1)
					{
						shapeRange.Top = shapeRange[1].Top;
						goto IL_0152;
					}
					rectangleF = A(application.ActivePresentation);
					if (rectangleF.HasValue)
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
						if (list[0] != rectangleF.Value.Top)
						{
							shapeRange.Top = rectangleF.Value.Top;
							goto IL_014a;
						}
					}
					shapeRange.Top = 0f;
					goto IL_014a;
				}
				RectangleF? rectangleF2 = A(application.ActivePresentation);
				if (rectangleF2.HasValue && shapeRange.Top != rectangleF2.Value.Top)
				{
					shapeRange.Top = rectangleF2.Value.Top;
				}
				else
				{
					shapeRange.Top = 0f;
				}
				rectangleF2 = null;
				goto IL_01b5;
				IL_014a:
				rectangleF = null;
				goto IL_0152;
				IL_0152:
				list = null;
				goto IL_01b5;
				IL_01b5:
				shapeRange = null;
				Base.LogActivity(AH.A(77790));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Base.AlignError();
				ProjectData.ClearProjectError();
			}
			application = null;
			return;
		}
	}

	internal static void K()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		D();
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		try
		{
			Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = Base.SelectedShapes();
			application.StartNewUndoEntry();
			if (shapeRange.Count > 1)
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
				List<float> list = new List<float>();
				{
					IEnumerator enumerator = shapeRange.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
							list.Add((float)Math.Round(shape.Top + shape.Height, 4));
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0099;
							}
							continue;
							end_IL_0099:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
				list = list.Distinct().ToList();
				float num;
				if (list.Count > 1)
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
					num = shapeRange[1].Top + shapeRange[1].Height;
				}
				else
				{
					Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
					RectangleF? rectangleF = A(activePresentation);
					num = ((!rectangleF.HasValue || (double)list[0] == Math.Round(rectangleF.Value.Top + rectangleF.Value.Height, 4)) ? activePresentation.PageSetup.SlideHeight : (rectangleF.Value.Top + rectangleF.Value.Height));
					rectangleF = null;
					activePresentation = null;
				}
				IEnumerator enumerator2 = default(IEnumerator);
				try
				{
					enumerator2 = shapeRange.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
						shape2.Top = num - shape2.Height;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
				list = null;
			}
			else
			{
				Microsoft.Office.Interop.PowerPoint.Presentation activePresentation2 = application.ActivePresentation;
				RectangleF? rectangleF2 = A(activePresentation2);
				Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange2 = shapeRange;
				if (rectangleF2.HasValue && Math.Round(shapeRange2.Top, 4) != Math.Round(rectangleF2.Value.Top + rectangleF2.Value.Height - shapeRange2.Height, 4))
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
					shapeRange2.Top = rectangleF2.Value.Top + rectangleF2.Value.Height - shapeRange2.Height;
				}
				else
				{
					shapeRange2.Top = activePresentation2.PageSetup.SlideHeight - shapeRange2.Height;
				}
				shapeRange2 = null;
				rectangleF2 = null;
				activePresentation2 = null;
			}
			Base.LogActivity(AH.A(78714));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Base.AlignError();
			ProjectData.ClearProjectError();
		}
		application = null;
	}

	internal static void L()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			E();
			Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
			try
			{
				Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = Base.SelectedShapes();
				application.StartNewUndoEntry();
				if (shapeRange.Count > 1)
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
					float num = shapeRange[1].Left + shapeRange[1].Width / 2f;
					try
					{
						enumerator = shapeRange.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
							shape.Left = num - shape.Width / 2f;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_00bb;
							}
							continue;
							end_IL_00bb:
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
				else
				{
					shapeRange[1].Left = application.ActivePresentation.PageSetup.SlideWidth / 2f - shapeRange[1].Width / 2f;
				}
				Base.LogActivity(AH.A(80533));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Base.AlignError();
				ProjectData.ClearProjectError();
			}
			application = null;
			return;
		}
	}

	internal static void M()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		F();
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		try
		{
			Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = Base.SelectedShapes();
			application.StartNewUndoEntry();
			if (shapeRange.Count > 1)
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
				float num = shapeRange[1].Top + shapeRange[1].Height / 2f;
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = shapeRange.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
						shape.Top = num - shape.Height / 2f;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_00b7;
						}
						continue;
						end_IL_00b7:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (7)
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
			else
			{
				shapeRange[1].Top = application.ActivePresentation.PageSetup.SlideHeight / 2f - shapeRange[1].Height / 2f;
			}
			Base.LogActivity(AH.A(80558));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Base.AlignError();
			ProjectData.ClearProjectError();
		}
		application = null;
	}

	private static RectangleF? A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		Settings settings = new Settings(A);
		RectangleF? result;
		if (settings.SlideMargins.HasValue)
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
			Settings.Margins value = settings.SlideMargins.Value;
			result = new RectangleF(value.Left, value.Top, A.PageSetup.SlideWidth - value.Left - value.Right, A.PageSetup.SlideHeight - value.Top - value.Bottom);
		}
		else
		{
			Microsoft.Office.Interop.PowerPoint.Shape bodyPlaceholder = Helpers.GetBodyPlaceholder(A);
			if (bodyPlaceholder != null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					try
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape = bodyPlaceholder;
						return new RectangleF(shape.Left, shape.Top, shape.Width, shape.Height);
					}
					finally
					{
						bodyPlaceholder = null;
					}
				}
			}
			result = null;
		}
		return result;
	}

	public static void AutoAlign()
	{
		if (!Licensing.AllowAdvancedShapeOperation())
		{
			return;
		}
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Slide slide = null;
		Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = null;
		List<Microsoft.Office.Interop.PowerPoint.Shape> list = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
		Selection selection;
		try
		{
			selection = application.ActiveWindow.Selection;
			if (selection.Type == PpSelectionType.ppSelectionShapes)
			{
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
					shapeRange = selection.ShapeRange;
					break;
				}
			}
			else
			{
				try
				{
					if (selection.SlideRange.Count == 1)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							if (Forms.OkCancelMessage2(AH.A(80583)) != DialogResult.OK)
							{
								break;
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								slide = selection.SlideRange[1];
								shapeRange = slide.Shapes.Range(RuntimeHelpers.GetObjectValue(Missing.Value));
								break;
							}
							break;
						}
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		List<int> list2;
		if (shapeRange != null)
		{
			if (shapeRange.Count >= 3)
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
				application.StartNewUndoEntry();
				int tolerance = GetTolerance();
				list = GetComparisonShapes(shapeRange);
				list2 = new List<int>();
				using (List<Microsoft.Office.Interop.PowerPoint.Shape>.Enumerator enumerator = list.GetEnumerator())
				{
					while (enumerator.MoveNext())
					{
						Microsoft.Office.Interop.PowerPoint.Shape current = enumerator.Current;
						Microsoft.Office.Interop.PowerPoint.Shape shape = current;
						MsoShapeType type = shape.Type;
						if (type == MsoShapeType.msoPlaceholder)
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
						bool flag = false;
						bool flag2 = false;
						int item = ((slide == null) ? shape.ZOrderPosition : Helpers.A(slide, current));
						double pos = Math.Round(shape.Top, MISALIGNMENT_ROUND);
						if (IsTopMisaligned(list, ref pos, tolerance))
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
							shape.Top = (float)pos;
							list2.Add(item);
							flag = true;
						}
						pos = Math.Round(shape.Left, MISALIGNMENT_ROUND);
						if (IsLeftMisaligned(list, ref pos, tolerance))
						{
							shape.Left = (float)pos;
							list2.Add(item);
							flag2 = true;
						}
						if (!flag)
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
							pos = Math.Round(shape.Top + shape.Height, MISALIGNMENT_ROUND);
							if (IsBottomMisaligned(list, ref pos, tolerance))
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
								shape.Top = (float)(pos - (double)shape.Height);
								list2.Add(item);
							}
						}
						if (!flag2)
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
							pos = Math.Round(shape.Left + shape.Width, MISALIGNMENT_ROUND);
							if (IsRightMisaligned(list, ref pos, tolerance))
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
								shape.Left = (float)(pos - (double)shape.Width);
								list2.Add(item);
							}
						}
						shape = null;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_02d5;
						}
						continue;
						end_IL_02d5:
						break;
					}
				}
				slide?.Shapes.Range(list2.Distinct().ToArray()).Select();
				int num = list2.Distinct().Count();
				if (num != 0)
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
					if (num == 1)
					{
						Forms.InfoMessage(AH.A(80774));
					}
					else
					{
						Forms.InfoMessage(AH.A(80829) + num + AH.A(80846));
					}
				}
				else
				{
					Forms.InfoMessage(AH.A(80719));
				}
				Base.LogActivity(AH.A(80885));
			}
			else
			{
				Forms.WarningMessage(AH.A(80906));
			}
			shapeRange = null;
		}
		application = null;
		selection = null;
		slide = null;
		list = null;
		list2 = null;
	}

	public static int GetTolerance()
	{
		return Conversions.ToInteger(KG.A.SettingsXml.DocumentElement.SelectSingleNode(AH.A(81009)).InnerText);
	}

	public static List<Microsoft.Office.Interop.PowerPoint.Shape> GetComparisonShapes(Microsoft.Office.Interop.PowerPoint.ShapeRange shapes)
	{
		List<Microsoft.Office.Interop.PowerPoint.Shape> list = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
		IEnumerator enumerator = shapes.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
				if (shape2.Visible == MsoTriState.msoTrue)
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
					MsoShapeType type = shape2.Type;
					if (type != MsoShapeType.msoFormControl)
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
						if (type != MsoShapeType.msoScriptAnchor)
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
							if ((uint)(type - 22) > 1u)
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
								object instance = NewLateBinding.LateGet(NewLateBinding.LateGet(shapes.Application, null, AH.A(81052), new object[0], null, null, null), null, AH.A(81089), new object[0], null, null, null);
								float num = Conversions.ToSingle(NewLateBinding.LateGet(instance, null, AH.A(81108), new object[0], null, null, null));
								float num2 = Conversions.ToSingle(NewLateBinding.LateGet(instance, null, AH.A(81131), new object[0], null, null, null));
								_ = null;
								if (shape2.Top + shape2.Height > 0f)
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
									if (shape2.Left + shape2.Width > 0f)
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
										if (shape2.Top < num)
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
											if (shape2.Left < num2)
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
												list.Add(shape);
											}
										}
									}
								}
							}
						}
					}
				}
				shape2 = null;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				return list;
			}
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	public static bool IsTopMisaligned(List<Microsoft.Office.Interop.PowerPoint.Shape> listShapes, ref double pos, int intTolerance)
	{
		double A = pos;
		List<Microsoft.Office.Interop.PowerPoint.Shape> list = listShapes.Where([SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape shape) => Align.A(A, Math.Round(shape.Top, MISALIGNMENT_ROUND), intTolerance)).ToList();
		if (Align.A(list))
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
			List<Microsoft.Office.Interop.PowerPoint.Shape> source = list;
			Func<Microsoft.Office.Interop.PowerPoint.Shape, double> keySelector;
			if (_Closure_0024__.A == null)
			{
				keySelector = (_Closure_0024__.A = [SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape shape) => Math.Round(shape.Top, MISALIGNMENT_ROUND));
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
				keySelector = _Closure_0024__.A;
			}
			Func<Microsoft.Office.Interop.PowerPoint.Shape, Microsoft.Office.Interop.PowerPoint.Shape> elementSelector = [SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape result) => result;
			Func<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>, BB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>> resultSelector;
			if (_Closure_0024__.A == null)
			{
				resultSelector = (_Closure_0024__.A = [SpecialName] (double a, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape> B) => new BB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>(a, B));
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
				resultSelector = _Closure_0024__.A;
			}
			IEnumerable<BB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>> source2 = source.GroupBy(keySelector, elementSelector, resultSelector);
			Func<BB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>, int> keySelector2;
			if (_Closure_0024__.A == null)
			{
				keySelector2 = (_Closure_0024__.A = [SpecialName] (BB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> bB) => bB.Group.Count());
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
				keySelector2 = _Closure_0024__.A;
			}
			IOrderedEnumerable<BB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>> source3 = source2.OrderByDescending(keySelector2);
			Func<BB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>, CB<double, int>> selector;
			if (_Closure_0024__.A == null)
			{
				selector = (_Closure_0024__.A = [SpecialName] (BB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> bB) => new CB<double, int>(bB.top, bB.Group.Count()));
			}
			else
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
				selector = _Closure_0024__.A;
			}
			IEnumerable<CB<double, int>> source4 = source3.Select(selector);
			if (source4.Count() >= 2 && source4.ElementAtOrDefault(0).cnt > source4.ElementAtOrDefault(1).cnt && pos != source4.ElementAtOrDefault(0).t)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						pos = source4.ElementAtOrDefault(0).t;
						list = null;
						return true;
					}
				}
			}
		}
		list = null;
		return false;
	}

	public static bool IsLeftMisaligned(List<Microsoft.Office.Interop.PowerPoint.Shape> listShapes, ref double pos, int intTolerance)
	{
		double A = pos;
		List<Microsoft.Office.Interop.PowerPoint.Shape> list = listShapes.Where([SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape shape) => Align.A(A, Math.Round(shape.Left, MISALIGNMENT_ROUND), intTolerance)).ToList();
		if (Align.A(list))
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
			List<Microsoft.Office.Interop.PowerPoint.Shape> source = list;
			Func<Microsoft.Office.Interop.PowerPoint.Shape, double> keySelector;
			if (_Closure_0024__.B == null)
			{
				keySelector = (_Closure_0024__.B = [SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape shape) => Math.Round(shape.Left, MISALIGNMENT_ROUND));
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
				keySelector = _Closure_0024__.B;
			}
			Func<Microsoft.Office.Interop.PowerPoint.Shape, Microsoft.Office.Interop.PowerPoint.Shape> elementSelector;
			if (_Closure_0024__.B == null)
			{
				elementSelector = (_Closure_0024__.B = [SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape result) => result);
			}
			else
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
				elementSelector = _Closure_0024__.B;
			}
			Func<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>, DB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>> resultSelector;
			if (_Closure_0024__.A == null)
			{
				resultSelector = (_Closure_0024__.A = [SpecialName] (double a, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape> B) => new DB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>(a, B));
			}
			else
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
				resultSelector = _Closure_0024__.A;
			}
			IEnumerable<DB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>> source2 = source.GroupBy(keySelector, elementSelector, resultSelector);
			Func<DB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>, int> keySelector2;
			if (_Closure_0024__.A == null)
			{
				keySelector2 = (_Closure_0024__.A = [SpecialName] (DB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> dB) => dB.Group.Count());
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
				keySelector2 = _Closure_0024__.A;
			}
			IEnumerable<EB<double, int>> source3 = from dB in source2.OrderByDescending(keySelector2)
				select new EB<double, int>(dB.left, dB.Group.Count());
			if (source3.Count() >= 2)
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
				if (source3.ElementAtOrDefault(0).cnt > source3.ElementAtOrDefault(1).cnt)
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
					if (pos != source3.ElementAtOrDefault(0).l)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								pos = source3.ElementAtOrDefault(0).l;
								list = null;
								return true;
							}
						}
					}
				}
			}
		}
		list = null;
		return false;
	}

	public static bool IsBottomMisaligned(List<Microsoft.Office.Interop.PowerPoint.Shape> listShapes, ref double pos, int intTolerance)
	{
		double A = pos;
		List<Microsoft.Office.Interop.PowerPoint.Shape> list = listShapes.Where([SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape shape) => Align.A(A, Math.Round(shape.Top + shape.Height, MISALIGNMENT_ROUND), intTolerance)).ToList();
		if (Align.A(list))
		{
			List<Microsoft.Office.Interop.PowerPoint.Shape> source = list;
			Func<Microsoft.Office.Interop.PowerPoint.Shape, double> keySelector = [SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape shape) => Math.Round(shape.Top + shape.Height, MISALIGNMENT_ROUND);
			Func<Microsoft.Office.Interop.PowerPoint.Shape, Microsoft.Office.Interop.PowerPoint.Shape> elementSelector;
			if (_Closure_0024__.C == null)
			{
				elementSelector = (_Closure_0024__.C = [SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape result) => result);
			}
			else
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
				elementSelector = _Closure_0024__.C;
			}
			Func<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>, FB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>> resultSelector;
			if (_Closure_0024__.A == null)
			{
				resultSelector = (_Closure_0024__.A = [SpecialName] (double a, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape> B) => new FB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>(a, B));
			}
			else
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
				resultSelector = _Closure_0024__.A;
			}
			IEnumerable<FB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>> source2 = source.GroupBy(keySelector, elementSelector, resultSelector);
			Func<FB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>, int> keySelector2;
			if (_Closure_0024__.A == null)
			{
				keySelector2 = (_Closure_0024__.A = [SpecialName] (FB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> fB) => fB.Group.Count());
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
				keySelector2 = _Closure_0024__.A;
			}
			IOrderedEnumerable<FB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>> source3 = source2.OrderByDescending(keySelector2);
			Func<FB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>, GB<double, int>> selector;
			if (_Closure_0024__.A == null)
			{
				selector = (_Closure_0024__.A = [SpecialName] (FB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>> fB) => new GB<double, int>(fB.bottom, fB.Group.Count()));
			}
			else
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
				selector = _Closure_0024__.A;
			}
			IEnumerable<GB<double, int>> source4 = source3.Select(selector);
			if (source4.Count() >= 2)
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
				if (source4.ElementAtOrDefault(0).cnt > source4.ElementAtOrDefault(1).cnt)
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
					if (pos != source4.ElementAtOrDefault(0).b)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								pos = source4.ElementAtOrDefault(0).b;
								list = null;
								return true;
							}
						}
					}
				}
			}
		}
		list = null;
		return false;
	}

	public static bool IsRightMisaligned(List<Microsoft.Office.Interop.PowerPoint.Shape> listShapes, ref double pos, int intTolerance)
	{
		double A = pos;
		List<Microsoft.Office.Interop.PowerPoint.Shape> list = listShapes.Where([SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape shape) => Align.A(A, Math.Round(shape.Left + shape.Width, MISALIGNMENT_ROUND), intTolerance)).ToList();
		if (Align.A(list))
		{
			List<Microsoft.Office.Interop.PowerPoint.Shape> source = list;
			Func<Microsoft.Office.Interop.PowerPoint.Shape, double> keySelector;
			if (_Closure_0024__.D == null)
			{
				keySelector = (_Closure_0024__.D = [SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape shape) => Math.Round(shape.Left + shape.Width, MISALIGNMENT_ROUND));
			}
			else
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
				keySelector = _Closure_0024__.D;
			}
			IEnumerable<IB<double, int>> source2 = from hB in source.GroupBy(keySelector, [SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape result) => result, [SpecialName] (double a, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape> B) => new HB<double, IEnumerable<Microsoft.Office.Interop.PowerPoint.Shape>>(a, B))
				orderby hB.Group.Count() descending
				select new IB<double, int>(hB.right, hB.Group.Count());
			if (source2.Count() >= 2)
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
				if (source2.ElementAtOrDefault(0).cnt > source2.ElementAtOrDefault(1).cnt)
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
					if (pos != source2.ElementAtOrDefault(0).r)
					{
						pos = source2.ElementAtOrDefault(0).r;
						list = null;
						return true;
					}
				}
			}
		}
		list = null;
		return false;
	}

	private static bool A(double A, double B, int C)
	{
		if (B > A - (double)C)
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
					return B < A + (double)C;
				}
			}
		}
		return false;
	}

	private static bool A(List<Microsoft.Office.Interop.PowerPoint.Shape> A)
	{
		return A.Count >= 3;
	}

	public static void OverTable()
	{
		if (!Licensing.AllowAdvancedShapeOperation())
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
			Table table = null;
			List<Microsoft.Office.Interop.PowerPoint.Shape> list;
			try
			{
				Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = Base.SelectedShapes();
				if (shapeRange.Count <= 1)
				{
					throw new Exception();
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					list = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
					try
					{
						foreach (Microsoft.Office.Interop.PowerPoint.Shape item in shapeRange)
						{
							if (item.HasTable == MsoTriState.msoTrue)
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
								if (table != null)
								{
									Forms.WarningMessage(AH.A(81152));
									table = null;
									break;
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
								table = item.Table;
							}
							else
							{
								list.Add(item);
							}
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					if (table == null)
					{
						break;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						if (list.Count <= 0)
						{
							break;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							NG.A.Application.StartNewUndoEntry();
							new wpfAlignOverTable(table, list).Show();
							_ = null;
							Base.LogActivity(AH.A(81211));
							break;
						}
						break;
					}
					break;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Helpers.TwoOrMoreShapesRequiredError();
				ProjectData.ClearProjectError();
			}
			table = null;
			list = null;
			return;
		}
	}
}
