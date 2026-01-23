using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Markup;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Utilities;

namespace PowerPointAddIn1.Shapes;

[DesignerGenerated]
public sealed class wpfAlignOverTable : Window, IComponentConnector, IStyleConnector
{
	public enum AlignPosition
	{
		TopLeft,
		TopCenter,
		TopRight,
		CenterLeft,
		CenterCenter,
		CenterRight,
		BottomLeft,
		BottomCenter,
		BottomRight
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Shape, float> A;

		public static Func<Shape, float> B;

		public static Func<ShapeCellDistance, float> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal float A(Shape A)
		{
			return A.Left;
		}

		[SpecialName]
		internal float B(Shape A)
		{
			return A.Top;
		}

		[SpecialName]
		internal float A(ShapeCellDistance A)
		{
			return A.Distance;
		}
	}

	private bool m_A;

	[CompilerGenerated]
	private Table m_A;

	[CompilerGenerated]
	private List<Shape> m_A;

	[CompilerGenerated]
	private Dictionary<Shape, System.Windows.Point> m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("radTopLeft")]
	private RadioButton m_A;

	[AccessedThroughProperty("radTopCenter")]
	[CompilerGenerated]
	private RadioButton m_B;

	[AccessedThroughProperty("radTopRight")]
	[CompilerGenerated]
	private RadioButton m_C;

	[AccessedThroughProperty("radCenterLeft")]
	[CompilerGenerated]
	private RadioButton m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("radCenterCenter")]
	private RadioButton m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("radCenterRight")]
	private RadioButton m_F;

	[CompilerGenerated]
	[AccessedThroughProperty("radBottomLeft")]
	private RadioButton m_G;

	[CompilerGenerated]
	[AccessedThroughProperty("radBottomCenter")]
	private RadioButton m_H;

	[AccessedThroughProperty("radBottomRight")]
	[CompilerGenerated]
	private RadioButton m_I;

	private bool m_B;

	private Table Table
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private List<Shape> Shapes
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private Dictionary<Shape, System.Windows.Point> OriginalLocations
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual RadioButton radTopLeft
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual RadioButton radTopCenter
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal virtual RadioButton radTopRight
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	internal virtual RadioButton radCenterLeft
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	internal virtual RadioButton radCenterCenter
	{
		[CompilerGenerated]
		get
		{
			return this.m_E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_E = value;
		}
	}

	internal virtual RadioButton radCenterRight
	{
		[CompilerGenerated]
		get
		{
			return this.m_F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_F = value;
		}
	}

	internal virtual RadioButton radBottomLeft
	{
		[CompilerGenerated]
		get
		{
			return this.m_G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_G = value;
		}
	}

	internal virtual RadioButton radBottomCenter
	{
		[CompilerGenerated]
		get
		{
			return this.m_H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_H = value;
		}
	}

	internal virtual RadioButton radBottomRight
	{
		[CompilerGenerated]
		get
		{
			return this.m_I;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_I = value;
		}
	}

	public wpfAlignOverTable(Table tbl, List<Shape> shps)
	{
		base.Loaded += wpfAlignOverTable_Loaded;
		base.Closing += wpfAlignOverTable_Closing;
		this.m_A = false;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		Table = tbl;
		Shapes = shps;
		OriginalLocations = new Dictionary<Shape, System.Windows.Point>();
		foreach (Shape shp in shps)
		{
			OriginalLocations.Add(shp, new System.Windows.Point(shp.Left, shp.Top));
		}
		radCenterCenter.IsChecked = true;
	}

	private void wpfAlignOverTable_Loaded(object sender, RoutedEventArgs e)
	{
		Dialogs.GetHwndSource(this).AddHook(A);
		A();
	}

	private void wpfAlignOverTable_Closing(object sender, CancelEventArgs e)
	{
		Dialogs.GetHwndSource(this).RemoveHook(A);
		Table = null;
		Shapes = null;
		OriginalLocations = null;
	}

	private void OnCloseExecuted(object sender, ExecutedRoutedEventArgs e)
	{
		Close();
	}

	private void A()
	{
		double C = default(double);
		double D = default(double);
		Dialogs.A((Shape)Table.Parent, Dialogs.DialogPosition.Left, ref C, ref D);
		base.Top = D;
		base.Left = C - base.Width;
	}

	private IntPtr A(IntPtr A, int B, IntPtr C, IntPtr D, ref bool E)
	{
		if (B == Dialogs.A)
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
			if (C.ToInt32() != 0)
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
				Show();
			}
			else
			{
				Hide();
				E = true;
			}
		}
		return IntPtr.Zero;
	}

	private void AlignmentChecked(object sender, RoutedEventArgs e)
	{
		if (sender == radCenterCenter)
		{
			A(AlignPosition.CenterCenter);
			return;
		}
		if (sender == radTopLeft)
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
					A(AlignPosition.TopLeft);
					return;
				}
			}
		}
		if (sender == radTopCenter)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					A(AlignPosition.TopCenter);
					return;
				}
			}
		}
		if (sender == radTopRight)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					A(AlignPosition.TopRight);
					return;
				}
			}
		}
		if (sender == radCenterLeft)
		{
			A(AlignPosition.CenterLeft);
			return;
		}
		if (sender == radCenterRight)
		{
			A(AlignPosition.CenterRight);
			return;
		}
		if (sender == radBottomLeft)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					A(AlignPosition.BottomLeft);
					return;
				}
			}
		}
		if (sender == radBottomCenter)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					A(AlignPosition.BottomCenter);
					return;
				}
			}
		}
		A(AlignPosition.BottomRight);
	}

	private void B()
	{
		using List<Shape>.Enumerator enumerator = Shapes.GetEnumerator();
		while (enumerator.MoveNext())
		{
			Shape current = enumerator.Current;
			System.Windows.Point point = OriginalLocations[current];
			current.Left = (float)point.X;
			current.Top = (float)point.Y;
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
			return;
		}
	}

	private void A(AlignPosition A)
	{
		if (this.m_A)
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
			try
			{
				B();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(Window.GetWindow(this), ex2.Message);
				Close();
				ProjectData.ClearProjectError();
				return;
			}
		}
		this.m_A = true;
		List<Shape> shapes = Shapes;
		Func<Shape, float> keySelector;
		if (_Closure_0024__.A == null)
		{
			keySelector = (_Closure_0024__.A = [SpecialName] (Shape shape) => shape.Left);
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
			keySelector = _Closure_0024__.A;
		}
		IOrderedEnumerable<Shape> source = shapes.OrderBy(keySelector);
		Func<Shape, float> keySelector2;
		if (_Closure_0024__.B == null)
		{
			keySelector2 = (_Closure_0024__.B = [SpecialName] (Shape shape) => shape.Top);
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
			keySelector2 = _Closure_0024__.B;
		}
		List<Shape> list = source.ThenBy(keySelector2).ToList();
		List<Shape> list2 = new List<Shape>();
		List<Cell> list3 = new List<Cell>();
		checked
		{
			using (List<Shape>.Enumerator enumerator = list.GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					Shape current = enumerator.Current;
					float num = 0f;
					Cell cell = null;
					RectangleF rectangleF = this.A(current);
					int count = Table.Rows.Count;
					for (int num2 = 1; num2 <= count; num2++)
					{
						int count2 = Table.Columns.Count;
						for (int num3 = 1; num3 <= count2; num3++)
						{
							Cell cell2 = Table.Cell(num2, num3);
							if (list3.Contains(cell2))
							{
								continue;
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
							RectangleF rect = this.A(cell2.Shape);
							RectangleF rectangleF2 = rectangleF;
							rectangleF2.Intersect(rect);
							if (rectangleF2 != Rectangle.Empty)
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
								float num4 = rectangleF2.Width * rectangleF2.Height;
								if (num4 > num)
								{
									num = num4;
									cell = cell2;
								}
							}
							rectangleF2 = default(RectangleF);
							cell2 = null;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_01e2;
							}
							continue;
							end_IL_01e2:
							break;
						}
					}
					if (cell != null)
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
						this.A(current, cell, A);
						list3.Add(cell);
					}
					else
					{
						list2.Add(current);
					}
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_0234;
					}
					continue;
					end_IL_0234:
					break;
				}
			}
			if (list2.Count <= 0)
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
				List<ShapeCellDistance> list4 = new List<ShapeCellDistance>();
				using (List<Shape>.Enumerator enumerator2 = list2.GetEnumerator())
				{
					while (enumerator2.MoveNext())
					{
						Shape current2 = enumerator2.Current;
						int count3 = Table.Rows.Count;
						for (int num5 = 1; num5 <= count3; num5++)
						{
							int count4 = Table.Columns.Count;
							for (int num6 = 1; num6 <= count4; num6++)
							{
								Cell cell2 = Table.Cell(num5, num6);
								if (list3.Contains(cell2))
								{
									continue;
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
								list4.Add(new ShapeCellDistance(current2, cell2, this.A(current2, cell2.Shape, A)));
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_0315;
								}
								continue;
								end_IL_0315:
								break;
							}
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
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0346;
						}
						continue;
						end_IL_0346:
						break;
					}
				}
				List<Shape> list5 = new List<Shape>();
				List<ShapeCellDistance> source2 = list4;
				Func<ShapeCellDistance, float> keySelector3;
				if (_Closure_0024__.A == null)
				{
					keySelector3 = (_Closure_0024__.A = [SpecialName] (ShapeCellDistance shapeCellDistance) => shapeCellDistance.Distance);
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
					keySelector3 = _Closure_0024__.A;
				}
				list4 = source2.OrderBy(keySelector3).ToList();
				using (List<ShapeCellDistance>.Enumerator enumerator3 = list4.GetEnumerator())
				{
					while (enumerator3.MoveNext())
					{
						ShapeCellDistance current3 = enumerator3.Current;
						if (list5.Contains(current3.Shape) || list3.Contains(current3.Cell))
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
						this.A(current3.Shape, current3.Cell, A);
						list5.Add(current3.Shape);
						list3.Add(current3.Cell);
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_042b;
						}
						continue;
						end_IL_042b:
						break;
					}
				}
				list5 = null;
				list4 = null;
				return;
			}
		}
	}

	private RectangleF A(Shape A)
	{
		return new RectangleF(A.Left, A.Top, A.Width, A.Height);
	}

	private float A(Shape A, Shape B, AlignPosition C)
	{
		System.Windows.Point point = default(System.Windows.Point);
		System.Windows.Point point2 = default(System.Windows.Point);
		switch (C)
		{
		case AlignPosition.CenterCenter:
			point = E(A);
			point2 = E(B);
			break;
		case AlignPosition.CenterLeft:
			point = D(A);
			point2 = D(B);
			break;
		case AlignPosition.CenterRight:
			point = F(A);
			point2 = F(B);
			break;
		case AlignPosition.TopLeft:
			point = this.A(A);
			point2 = this.A(B);
			break;
		case AlignPosition.TopCenter:
			point = this.B(A);
			point2 = this.B(B);
			break;
		case AlignPosition.TopRight:
			point = this.C(A);
			point2 = this.C(B);
			break;
		case AlignPosition.BottomLeft:
			point = G(A);
			point2 = G(B);
			break;
		case AlignPosition.BottomCenter:
			point = H(A);
			point2 = H(B);
			break;
		case AlignPosition.BottomRight:
			point = I(A);
			point2 = I(B);
			break;
		}
		return (float)Math.Sqrt(Math.Pow(point2.X - point.X, 2.0) + Math.Pow(point2.Y - point.Y, 2.0));
	}

	private System.Windows.Point A(Shape A)
	{
		return new System.Windows.Point(A.Left, A.Top);
	}

	private System.Windows.Point B(Shape A)
	{
		return new System.Windows.Point(A.Left + A.Width / 2f, A.Top);
	}

	private System.Windows.Point C(Shape A)
	{
		return new System.Windows.Point(A.Left + A.Width, A.Top);
	}

	private System.Windows.Point D(Shape A)
	{
		return new System.Windows.Point(A.Left, A.Top + A.Height / 2f);
	}

	private System.Windows.Point E(Shape A)
	{
		return new System.Windows.Point(A.Left + A.Width / 2f, A.Top + A.Height / 2f);
	}

	private System.Windows.Point F(Shape A)
	{
		return new System.Windows.Point(A.Left + A.Width, A.Top + A.Height / 2f);
	}

	private System.Windows.Point G(Shape A)
	{
		return new System.Windows.Point(A.Left, A.Top + A.Height);
	}

	private System.Windows.Point H(Shape A)
	{
		return new System.Windows.Point(A.Left + A.Width / 2f, A.Top + A.Height);
	}

	private System.Windows.Point I(Shape A)
	{
		return new System.Windows.Point(A.Left + A.Width, A.Top + A.Height);
	}

	private void A(Shape A, Cell B, AlignPosition C)
	{
		Shape shape = B.Shape;
		if ((uint)C > 2u)
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
			if ((uint)(C - 3) <= 2u)
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
				A.Top = shape.Top + (shape.Height - A.Height) / 2f;
			}
			else
			{
				A.Top = shape.Top + shape.Height - A.Height;
			}
		}
		else
		{
			A.Top = shape.Top;
		}
		switch (C)
		{
		case AlignPosition.TopCenter:
		case AlignPosition.CenterCenter:
		case AlignPosition.BottomCenter:
			A.Left = shape.Left + (shape.Width - A.Width) / 2f;
			break;
		case AlignPosition.TopLeft:
		case AlignPosition.CenterLeft:
		case AlignPosition.BottomLeft:
			A.Left = shape.Left;
			break;
		default:
			A.Left = shape.Left + shape.Width - A.Width;
			break;
		}
		shape = null;
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_B)
		{
			this.m_B = true;
			Uri resourceLocator = new Uri(AH.A(91966), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
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
					((CommandBinding)target).Executed += OnCloseExecuted;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					radTopLeft = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					radTopCenter = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			radTopRight = (RadioButton)target;
			return;
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					radCenterLeft = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			radCenterCenter = (RadioButton)target;
			return;
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					radCenterRight = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					radBottomLeft = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					radBottomCenter = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					radBottomRight = (RadioButton)target;
					return;
				}
			}
		}
		this.m_B = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 2)
		{
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = ToggleButton.CheckedEvent;
			eventSetter.Handler = new RoutedEventHandler(AlignmentChecked);
			((Style)target).Setters.Add(eventSetter);
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
