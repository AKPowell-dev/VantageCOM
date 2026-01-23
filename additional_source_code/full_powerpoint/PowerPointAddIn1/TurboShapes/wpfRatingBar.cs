using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using A;
using Foo.Controls;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Utilities;

namespace PowerPointAddIn1.TurboShapes;

[DesignerGenerated]
public sealed class wpfRatingBar : Window, IComponentConnector
{
	public Shape EditedShape;

	public int CurrentSegments;

	public List<RatingBar.BarStyle> BarStyles;

	public RatingBar.BarStyle CurrentBarStyle;

	private Microsoft.Office.Interop.PowerPoint.Application m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("Slider")]
	private Slider m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("numSegments")]
	private MacNumericUpDown m_A;

	[AccessedThroughProperty("cbxShapes")]
	[CompilerGenerated]
	private ComboBox m_A;

	[AccessedThroughProperty("btnStyle")]
	[CompilerGenerated]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClose")]
	private Button B;

	private bool m_A;

	internal virtual Slider Slider
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

	internal virtual MacNumericUpDown numSegments
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

	internal virtual ComboBox cbxShapes
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

	internal virtual Button btnStyle
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
			RoutedEventHandler value2 = ChangeStyle;
			Button button = this.m_A;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.m_A = value;
			button = this.m_A;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual Button btnClose
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnClose_Click;
			Button button = B;
			if (button != null)
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
				button.Click -= value2;
			}
			B = value;
			button = B;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	public wpfRatingBar()
	{
		base.Loaded += wpfRatingBar_Loaded;
		base.KeyDown += wpfHarveyBall_KeyDown;
		base.Closing += wpfRatingBar_Closing;
		InitializeComponent();
		this.m_A = NG.A.Application;
		Base.ActiveCallout = this;
	}

	private IntPtr A(IntPtr A, int B, IntPtr C, IntPtr D, ref bool E)
	{
		if (B == Dialogs.A)
		{
			if (C.ToInt32() != 0)
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

	private void wpfRatingBar_Loaded(object sender, RoutedEventArgs e)
	{
		//IL_003f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0049: Expected O, but got Unknown
		Dialogs.GetHwndSource(this).AddHook(A);
		Slider.ValueChanged += Slider_ValueChanged;
		numSegments.ValueChanged += new MacRangeBaseValueChangedHandler(numSegments_ValueChanged);
		cbxShapes.SelectionChanged += cbxShapes_SelectionChanged;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(113083)).AddEventHandler(this.m_A, new EApplication_AfterShapeSizeChangeEventHandler(A));
	}

	private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
	{
		A();
	}

	private void numSegments_ValueChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		int num = checked((int)Math.Round(numSegments.Value.Value));
		Slider slider = Slider;
		if (slider.Value > (double)num)
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
			slider.ValueChanged -= Slider_ValueChanged;
			slider.Value = num;
			slider.ValueChanged += Slider_ValueChanged;
		}
		slider.Maximum = num;
		slider = null;
		A();
	}

	private void cbxShapes_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		A();
	}

	private void ChangeStyle(object sender, RoutedEventArgs e)
	{
		int num = BarStyles.IndexOf(CurrentBarStyle);
		num = checked((num != BarStyles.Count - 1) ? (num + 1) : 0);
		CurrentBarStyle = BarStyles[num];
		A();
	}

	private void A()
	{
		Slide slide = Base.GetSlide(EditedShape);
		RatingBar.ShapeType type = default(RatingBar.ShapeType);
		switch (cbxShapes.SelectedIndex)
		{
		case 0:
			type = RatingBar.ShapeType.Circle;
			break;
		case 1:
			type = RatingBar.ShapeType.Square;
			break;
		case 2:
			type = RatingBar.ShapeType.Rectangle;
			break;
		case 3:
			type = RatingBar.ShapeType.Star;
			break;
		case 4:
			type = RatingBar.ShapeType.Diamond;
			break;
		}
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(A));
		Base.RemoveSelectionChangedEvent();
		base.Deactivated -= CalloutDeactivated;
		checked
		{
			Shape shape;
			try
			{
				this.m_A.StartNewUndoEntry();
				Shape editedShape = EditedShape;
				shape = RatingBar.Create(slide, (float)Slider.Value, (int)Math.Round(numSegments.Value.Value), type, CurrentBarStyle);
				Base.ImportPrimaryColor(shape, EditedShape);
				shape.Height = editedShape.Height;
				shape.Top = editedShape.Top;
				shape.Left = editedShape.Left;
				editedShape.Delete();
				editedShape = null;
				EditedShape = shape;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(ex2.Message);
				Base.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			CurrentSegments = (int)Math.Round(numSegments.Value.Value);
			Activate();
			base.Deactivated += CalloutDeactivated;
			Base.AddSelectionChangedEvent();
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(A));
			slide = null;
			shape = null;
		}
	}

	private void wpfHarveyBall_KeyDown(object sender, KeyEventArgs e)
	{
		Key key = e.Key;
		if (key != Key.Return)
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
			if (key != Key.Escape)
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
				break;
			}
		}
		Close();
	}

	private void btnClose_Click(object sender, RoutedEventArgs e)
	{
		Close();
	}

	private void A(Selection A)
	{
		if (base.IsActive)
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
			Close();
			return;
		}
	}

	private void CalloutDeactivated(object sender, EventArgs e)
	{
		Close();
	}

	private void A(Shape A)
	{
		double unitX = default(double);
		double unitY = default(double);
		Base.TransformFromShape(EditedShape, Base.CalloutPosition.TopLeft, ref unitX, ref unitY);
		base.Top = unitY - base.Height;
		base.Left = unitX;
	}

	private void wpfRatingBar_Closing(object sender, CancelEventArgs e)
	{
		Dialogs.GetHwndSource(this).RemoveHook(A);
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(113083)).RemoveEventHandler(this.m_A, new EApplication_AfterShapeSizeChangeEventHandler(A));
		base.Deactivated -= CalloutDeactivated;
		EditedShape = null;
		BarStyles = null;
		this.m_A = null;
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_A)
		{
			this.m_A = true;
			Uri resourceLocator = new Uri(AH.A(160095), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003e: Expected O, but got Unknown
		if (connectionId == 1)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Slider = (Slider)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					numSegments = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					cbxShapes = (ComboBox)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 4:
			btnStyle = (Button)target;
			break;
		case 5:
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				btnClose = (Button)target;
				return;
			}
		default:
			this.m_A = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
