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
public sealed class wpfSliderBar : Window, IComponentConnector
{
	public Microsoft.Office.Interop.PowerPoint.Shape EditedShape;

	public List<SliderBar.BarStyle> BarStyles;

	public SliderBar.BarStyle CurrentBarStyle;

	public SliderBar.SliderStyle CurrentSliderStyle;

	private Microsoft.Office.Interop.PowerPoint.Application m_A;

	[AccessedThroughProperty("Slider")]
	[CompilerGenerated]
	private Slider m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("numValue")]
	private MacNumericUpDown m_A;

	[AccessedThroughProperty("cbxSlider")]
	[CompilerGenerated]
	private ComboBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnLeft")]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnRight")]
	private Button m_B;

	[AccessedThroughProperty("btnClose")]
	[CompilerGenerated]
	private Button C;

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

	internal virtual ComboBox cbxSlider
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

	internal virtual Button btnLeft
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
			RoutedEventHandler value2 = btnLeft_Click;
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

	internal virtual Button btnRight
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
			RoutedEventHandler value2 = btnRight_Click;
			Button button = this.m_B;
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
			this.m_B = value;
			button = this.m_B;
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

	internal virtual Button btnClose
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
			RoutedEventHandler value2 = btnClose_Click;
			Button button = C;
			if (button != null)
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
				button.Click -= value2;
			}
			C = value;
			button = C;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	public wpfSliderBar()
	{
		base.Loaded += wpfHarveyBall_Loaded;
		base.KeyDown += wpfHarveyBall_KeyDown;
		base.Closing += wpfHarveyBall_Closing;
		InitializeComponent();
		Base.SetImageSource((Image)((StackPanel)((ComboBoxItem)cbxSlider.Items[1]).Content).Children[0], OB.TurboShapePentagon);
		Base.SetImageSource((Image)((StackPanel)((ComboBoxItem)cbxSlider.Items[2]).Content).Children[0], OB.TurboShapeRectangle);
		this.m_A = NG.A.Application;
		Base.ActiveCallout = this;
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
		//IL_0041: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		Dialogs.GetHwndSource(this).AddHook(A);
		Slider.ValueChanged += Slider_ValueChanged;
		numValue.ValueChanged += new MacRangeBaseValueChangedHandler(numValue_ValueChanged);
		cbxSlider.SelectionChanged += cbxSlider_SelectionChanged;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(113083)).AddEventHandler(this.m_A, new EApplication_AfterShapeSizeChangeEventHandler(A));
	}

	private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
	{
		checked
		{
			B((int)Math.Round(e.NewValue));
			A((int)Math.Round(e.NewValue));
		}
	}

	private void numValue_ValueChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		int num = checked((int)Math.Round(numValue.Value.Value));
		B(num);
		Slider slider = Slider;
		slider.ValueChanged -= Slider_ValueChanged;
		slider.Value = num;
		slider.ValueChanged += Slider_ValueChanged;
		_ = null;
	}

	private void A(int A)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		//IL_002c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0036: Expected O, but got Unknown
		MacNumericUpDown obj = numValue;
		obj.ValueChanged -= new MacRangeBaseValueChangedHandler(numValue_ValueChanged);
		obj.Value = A;
		obj.ValueChanged += new MacRangeBaseValueChangedHandler(numValue_ValueChanged);
		_ = null;
	}

	private void cbxSlider_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		SliderBar.SliderStyle currentSliderStyle = default(SliderBar.SliderStyle);
		switch (cbxSlider.SelectedIndex)
		{
		case 0:
			currentSliderStyle = SliderBar.SliderStyle.Circle;
			break;
		case 1:
			currentSliderStyle = SliderBar.SliderStyle.Pentagon;
			break;
		case 2:
			currentSliderStyle = SliderBar.SliderStyle.Rectangle;
			break;
		}
		CurrentSliderStyle = currentSliderStyle;
		B(checked((int)Math.Round(numValue.Value.Value)));
	}

	private void btnLeft_Click(object sender, RoutedEventArgs e)
	{
		int num = BarStyles.IndexOf(CurrentBarStyle);
		checked
		{
			if (num == 0)
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
				num = BarStyles.Count - 1;
			}
			else
			{
				num--;
			}
			CurrentBarStyle = BarStyles[num];
			B((int)Math.Round(numValue.Value.Value));
		}
	}

	private void btnRight_Click(object sender, RoutedEventArgs e)
	{
		int num = BarStyles.IndexOf(CurrentBarStyle);
		checked
		{
			if (num == BarStyles.Count - 1)
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
				num = 0;
			}
			else
			{
				num++;
			}
			CurrentBarStyle = BarStyles[num];
			B((int)Math.Round(numValue.Value.Value));
		}
	}

	private void B(int A)
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
			shape = SliderBar.Create(slide, A, CurrentBarStyle, CurrentSliderStyle);
			if (shape.Type == MsoShapeType.msoGroup)
			{
				if (editedShape.Type == MsoShapeType.msoGroup)
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
					shape.GroupItems[2].Fill.ForeColor.RGB = editedShape.GroupItems[2].Fill.ForeColor.RGB;
					shape.GroupItems[2].Fill.BackColor.RGB = editedShape.GroupItems[2].Fill.BackColor.RGB;
				}
				else
				{
					shape.GroupItems[2].Fill.ForeColor.RGB = editedShape.Fill.ForeColor.RGB;
					shape.GroupItems[2].Fill.BackColor.RGB = editedShape.Fill.BackColor.RGB;
				}
			}
			else if (editedShape.Type == MsoShapeType.msoGroup)
			{
				shape.Fill.ForeColor.RGB = editedShape.GroupItems[2].Fill.ForeColor.RGB;
				shape.Fill.BackColor.RGB = editedShape.GroupItems[2].Fill.BackColor.RGB;
			}
			else
			{
				shape.Fill.ForeColor.RGB = editedShape.Fill.ForeColor.RGB;
				shape.Fill.BackColor.RGB = editedShape.Fill.BackColor.RGB;
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
			switch (3)
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
		Base.TransformFromShape(EditedShape, Base.CalloutPosition.TopLeft, ref unitX, ref unitY);
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
			Uri resourceLocator = new Uri(AH.A(160509), UriKind.Relative);
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
					numValue = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					cbxSlider = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnLeft = (Button)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 5:
			btnRight = (Button)target;
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
