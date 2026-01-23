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
public sealed class wpfTachometer : Window, IComponentConnector
{
	public Shape EditedShape;

	public List<Tachometer.Style> Styles;

	public Tachometer.Style CurrentStyle;

	public bool IsReversed;

	public bool ShowLabel;

	private Microsoft.Office.Interop.PowerPoint.Application m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("Slider")]
	private Slider m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("numValue")]
	private MacNumericUpDown m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnFlip")]
	private Button m_A;

	[AccessedThroughProperty("btnLabel")]
	[CompilerGenerated]
	private Button B;

	[AccessedThroughProperty("btnStyle")]
	[CompilerGenerated]
	private Button C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClose")]
	private Button D;

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

	internal virtual MacNumericUpDown numValue
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

	internal virtual Button btnFlip
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
			RoutedEventHandler value2 = btnFlip_Click;
			Button button = this.m_A;
			if (button != null)
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

	internal virtual Button btnLabel
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
			RoutedEventHandler value2 = btnLabel_Click;
			Button button = B;
			if (button != null)
			{
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
				switch (3)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnStyle
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnStyle_Click;
			Button button = C;
			if (button != null)
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
				button.Click -= value2;
			}
			C = value;
			button = C;
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
			return D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnClose_Click;
			Button button = D;
			if (button != null)
			{
				button.Click -= value2;
			}
			D = value;
			button = D;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	public wpfTachometer()
	{
		base.Loaded += wpfTachometer_Loaded;
		base.KeyDown += wpfTachometer_KeyDown;
		base.Closing += wpfTachometer_Closing;
		InitializeComponent();
		this.m_A = NG.A.Application;
		Base.ActiveCallout = this;
	}

	private IntPtr A(IntPtr A, int B, IntPtr C, IntPtr D, ref bool E)
	{
		if (B == Dialogs.A)
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
			if (C.ToInt32() != 0)
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

	private void wpfTachometer_Loaded(object sender, RoutedEventArgs e)
	{
		//IL_0041: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		Dialogs.GetHwndSource(this).AddHook(A);
		Slider.ValueChanged += Slider_ValueChanged;
		numValue.ValueChanged += new MacRangeBaseValueChangedHandler(numValue_ValueChanged);
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(113083)).AddEventHandler(this.m_A, new EApplication_AfterShapeSizeChangeEventHandler(A));
	}

	private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
	{
		A(e.NewValue);
		A(checked((int)Math.Round(e.NewValue)));
	}

	private void numValue_ValueChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		int num = checked((int)Math.Round(numValue.Value.Value));
		A((double)num);
		Slider slider = Slider;
		slider.ValueChanged -= Slider_ValueChanged;
		slider.Value = num;
		slider.ValueChanged += Slider_ValueChanged;
		_ = null;
	}

	private void A(int A)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_001a: Expected O, but got Unknown
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		MacNumericUpDown obj = numValue;
		obj.ValueChanged -= new MacRangeBaseValueChangedHandler(numValue_ValueChanged);
		obj.Value = A;
		obj.ValueChanged += new MacRangeBaseValueChangedHandler(numValue_ValueChanged);
		_ = null;
	}

	private void btnFlip_Click(object sender, RoutedEventArgs e)
	{
		IsReversed = !IsReversed;
		A(numValue.Value.Value);
	}

	private void btnLabel_Click(object sender, RoutedEventArgs e)
	{
		ShowLabel = !ShowLabel;
		A(numValue.Value.Value);
	}

	private void btnStyle_Click(object sender, RoutedEventArgs e)
	{
		int num = Styles.IndexOf(CurrentStyle);
		num = checked((num != Styles.Count - 1) ? (num + 1) : 0);
		CurrentStyle = Styles[num];
		A(numValue.Value.Value);
	}

	private void A(double A)
	{
		Slide slide = Base.GetSlide(EditedShape);
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(this.A));
		Base.RemoveSelectionChangedEvent();
		base.Deactivated -= CalloutDeactivated;
		Shape shape;
		try
		{
			this.m_A.StartNewUndoEntry();
			Shape editedShape = EditedShape;
			shape = Tachometer.Create(slide, (float)A, (int)CurrentStyle, IsReversed, ShowLabel);
			shape.Width = editedShape.Width;
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
		Activate();
		base.Deactivated += CalloutDeactivated;
		Base.AddSelectionChangedEvent();
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(this.A));
		slide = null;
		shape = null;
	}

	private void wpfTachometer_KeyDown(object sender, KeyEventArgs e)
	{
		Key key = e.Key;
		if (key == Key.Escape)
		{
			Close();
		}
	}

	private void btnClose_Click(object sender, RoutedEventArgs e)
	{
		Close();
	}

	private void A(Selection A)
	{
		if (!base.IsActive)
		{
			Close();
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
		Base.TransformFromShape(EditedShape, Base.CalloutPosition.TopCenter, ref unitX, ref unitY);
		base.Top = unitY - base.Height;
		base.Left = unitX;
	}

	private void wpfTachometer_Closing(object sender, CancelEventArgs e)
	{
		Dialogs.GetHwndSource(this).RemoveHook(A);
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(113083)).RemoveEventHandler(this.m_A, new EApplication_AfterShapeSizeChangeEventHandler(A));
		base.Deactivated -= CalloutDeactivated;
		EditedShape = null;
		this.m_A = null;
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (this.m_A)
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
			this.m_A = true;
			Uri resourceLocator = new Uri(AH.A(160231), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
			return;
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
		//IL_002a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0034: Expected O, but got Unknown
		if (connectionId == 1)
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
					Slider = (Slider)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			numValue = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnFlip = (Button)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 4:
			btnLabel = (Button)target;
			break;
		case 5:
			btnStyle = (Button)target;
			break;
		case 6:
			while (true)
			{
				switch (4)
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
