using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using A;
using Foo.Controls;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Utilities;

namespace PowerPointAddIn1.Shapes;

[DesignerGenerated]
public sealed class wpfGapSize : Window, IComponentConnector
{
	[CompilerGenerated]
	private Shape m_A;

	[CompilerGenerated]
	private ShapeRange m_A;

	[CompilerGenerated]
	private Action<ShapeRange, float> m_A;

	[CompilerGenerated]
	private Dialogs.DialogPosition m_A;

	[AccessedThroughProperty("numGap")]
	[CompilerGenerated]
	private MacNumericUpDown m_A;

	private bool m_A;

	private Shape ReferenceShape
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

	private ShapeRange Shapes
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

	private Action<ShapeRange, float> MoveCallback
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

	private Dialogs.DialogPosition DialogPos
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

	internal virtual MacNumericUpDown numGap
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

	public wpfGapSize(Shape shpRef, Dialogs.DialogPosition pos, ShapeRange shps, Action<ShapeRange, float> act)
	{
		//IL_0092: Unknown result type (might be due to invalid IL or missing references)
		//IL_0097: Unknown result type (might be due to invalid IL or missing references)
		//IL_0099: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bc: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0112: Unknown result type (might be due to invalid IL or missing references)
		//IL_011c: Expected O, but got Unknown
		base.Loaded += wpfGapSize_Loaded;
		base.PreviewKeyDown += wpfGapSize_PreviewKeyDown;
		base.Deactivated += wpfGapSize_Deactivated;
		base.Closing += wpfGapSize_Closing;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		ReferenceShape = shpRef;
		Shapes = shps;
		MoveCallback = act;
		DialogPos = pos;
		bool isMetric = RegionInfo.CurrentRegion.IsMetric;
		SpinnerProperties spinnerProps = Regional.GetSpinnerProps((bool?)isMetric);
		MacNumericUpDown val = numGap;
		val.CustomUnit = spinnerProps.CustomUnits;
		val.NumberDecimalDigits = spinnerProps.Decimals;
		val.SmallChange = spinnerProps.SmallChange;
		val.LargeChange = spinnerProps.LargeChange;
		if (isMetric)
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
			val.Maximum = 20.0;
		}
		else
		{
			val.Maximum = 8.0;
		}
		val.ValueChanged += new MacRangeBaseValueChangedHandler(GapSizeChanged);
		val.Value = PB.Settings.LastGapSize;
		val = null;
	}

	private void wpfGapSize_Loaded(object sender, RoutedEventArgs e)
	{
		Dialogs.GetHwndSource(this).AddHook(A);
		A();
	}

	private void wpfGapSize_PreviewKeyDown(object sender, KeyEventArgs e)
	{
		if (e.IsRepeat)
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
			if (((UIElement)(object)numGap).IsKeyboardFocused)
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
				Key key = e.Key;
				if ((uint)(key - 19) > 1u && key != Key.Up)
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
					if (key != Key.Down)
					{
						return;
					}
				}
				((UIElement)(object)numGap).Focus();
				return;
			}
		}
	}

	private void wpfGapSize_Deactivated(object sender, EventArgs e)
	{
	}

	private void wpfGapSize_Closing(object sender, CancelEventArgs e)
	{
		//IL_003d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Expected O, but got Unknown
		Dialogs.GetHwndSource(this).RemoveHook(A);
		ReferenceShape = null;
		Shapes = null;
		MoveCallback = null;
		MacNumericUpDown val = numGap;
		val.ValueChanged -= new MacRangeBaseValueChangedHandler(GapSizeChanged);
		if (val.Value.HasValue)
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
			PB.Settings.LastGapSize = (float)val.Value.Value;
		}
		val = null;
	}

	private void CloseDialog(object sender, RoutedEventArgs e)
	{
		Close();
	}

	private void A()
	{
		Dialogs.DialogPosition dialogPos = DialogPos;
		double C = default(double);
		double D = default(double);
		Dialogs.A(ReferenceShape, dialogPos, ref C, ref D);
		if (dialogPos != Dialogs.DialogPosition.Above)
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
					if (dialogPos != Dialogs.DialogPosition.Left)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								base.Top = D;
								base.Left = C;
								return;
							}
						}
					}
					base.Top = D;
					base.Left = C - base.Width;
					return;
				}
			}
		}
		base.Top = D - base.Height;
		base.Left = C;
	}

	private IntPtr A(IntPtr A, int B, IntPtr C, IntPtr D, ref bool E)
	{
		if (B == Dialogs.A)
		{
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

	private void GapSizeChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		float a = 0f;
		if (numGap.Value.HasValue)
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
			a = (float)numGap.Value.Value;
		}
		MoveCallback(Shapes, A(a));
	}

	private float A(float A)
	{
		if (RegionInfo.CurrentRegion.IsMetric)
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
					return clsPublish.CentimetersToPoints(A);
				}
			}
		}
		return clsPublish.InchesToPoints(A);
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (!this.m_A)
		{
			this.m_A = true;
			Uri resourceLocator = new Uri(AH.A(75626), UriKind.Relative);
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
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0010: Expected O, but got Unknown
		switch (connectionId)
		{
		case 1:
			numGap = (MacNumericUpDown)target;
			break;
		case 2:
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
				((Button)target).Click += CloseDialog;
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
