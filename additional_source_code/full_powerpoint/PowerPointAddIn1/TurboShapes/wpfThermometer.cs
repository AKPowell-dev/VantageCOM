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
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Utilities;

namespace PowerPointAddIn1.TurboShapes;

[DesignerGenerated]
public sealed class wpfThermometer : Window, IComponentConnector
{
	public Microsoft.Office.Interop.PowerPoint.Shape EditedShape;

	public List<Thermometer.MeterStyle> MeterStyles;

	public Thermometer.MeterStyle CurrentStyle;

	private Microsoft.Office.Interop.PowerPoint.Application m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("Slider")]
	private Slider m_A;

	[AccessedThroughProperty("numValue")]
	[CompilerGenerated]
	private MacNumericUpDown m_A;

	[AccessedThroughProperty("btnStyle")]
	[CompilerGenerated]
	private Button m_A;

	[AccessedThroughProperty("btnClose")]
	[CompilerGenerated]
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
			RoutedEventHandler value2 = btnStyle_Click;
			Button button = this.m_A;
			if (button != null)
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
				button.Click -= value2;
			}
			this.m_A = value;
			button = this.m_A;
			if (button == null)
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
				button.Click += value2;
				return;
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
				switch (4)
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

	public wpfThermometer()
	{
		base.Loaded += wpfHarveyBall_Loaded;
		base.KeyDown += wpfHarveyBall_KeyDown;
		base.Closing += wpfHarveyBall_Closing;
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

	private void wpfHarveyBall_Loaded(object sender, RoutedEventArgs e)
	{
		//IL_003d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Expected O, but got Unknown
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

	private void btnStyle_Click(object sender, RoutedEventArgs e)
	{
		int num = MeterStyles.IndexOf(CurrentStyle);
		checked
		{
			if (num == MeterStyles.Count - 1)
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
				num = 0;
			}
			else
			{
				num++;
			}
			CurrentStyle = MeterStyles[num];
			A(numValue.Value.Value);
		}
	}

	private void A(double A)
	{
		Slide slide = Base.GetSlide(EditedShape);
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(this.A));
		Base.RemoveSelectionChangedEvent();
		base.Deactivated -= CalloutDeactivated;
		Microsoft.Office.Interop.PowerPoint.Shape shape;
		try
		{
			this.m_A.StartNewUndoEntry();
			Microsoft.Office.Interop.PowerPoint.Shape editedShape = EditedShape;
			shape = Thermometer.Create(slide, checked((int)Math.Round(A)), CurrentStyle);
			if (CurrentStyle == Thermometer.MeterStyle.OneColor)
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
				if (editedShape.Type == MsoShapeType.msoGroup)
				{
					shape.Fill.ForeColor.RGB = editedShape.GroupItems[2].Fill.ForeColor.RGB;
					shape.Fill.BackColor.RGB = editedShape.GroupItems[2].Fill.BackColor.RGB;
				}
				else
				{
					shape.Fill.ForeColor.RGB = editedShape.Fill.ForeColor.RGB;
					shape.Fill.BackColor.RGB = editedShape.Fill.BackColor.RGB;
				}
			}
			else if (editedShape.Type == MsoShapeType.msoGroup)
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
				shape.GroupItems[2].Fill.ForeColor.RGB = editedShape.GroupItems[2].Fill.ForeColor.RGB;
				shape.GroupItems[2].Fill.BackColor.RGB = editedShape.GroupItems[2].Fill.BackColor.RGB;
			}
			else
			{
				shape.GroupItems[2].Fill.ForeColor.RGB = editedShape.Fill.ForeColor.RGB;
				shape.GroupItems[2].Fill.BackColor.RGB = editedShape.Fill.BackColor.RGB;
			}
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
		Activate();
		base.Deactivated += CalloutDeactivated;
		Base.AddSelectionChangedEvent();
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(this.A));
		slide = null;
		shape = null;
	}

	private void wpfHarveyBall_KeyDown(object sender, KeyEventArgs e)
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
		if (base.IsActive)
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
			Close();
			return;
		}
	}

	private void CalloutDeactivated(object sender, EventArgs e)
	{
		Close();
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		double unitX = default(double);
		double unitY = default(double);
		Base.TransformFromShape(EditedShape, Base.CalloutPosition.TopCenter, ref unitX, ref unitY);
		base.Top = unitY - base.Height;
		base.Left = unitX;
	}

	private void wpfHarveyBall_Closing(object sender, CancelEventArgs e)
	{
		Dialogs.GetHwndSource(this).RemoveHook(A);
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(113083)).RemoveEventHandler(this.m_A, new EApplication_AfterShapeSizeChangeEventHandler(A));
		base.Deactivated -= CalloutDeactivated;
		EditedShape = null;
		MeterStyles = null;
		this.m_A = null;
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_A)
		{
			this.m_A = true;
			Uri resourceLocator = new Uri(AH.A(160369), UriKind.Relative);
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
		//IL_002a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0034: Expected O, but got Unknown
		if (connectionId == 1)
		{
			while (true)
			{
				switch (4)
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
				switch (4)
				{
				case 0:
					break;
				default:
					btnStyle = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnClose = (Button)target;
					return;
				}
			}
		}
		this.m_A = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
