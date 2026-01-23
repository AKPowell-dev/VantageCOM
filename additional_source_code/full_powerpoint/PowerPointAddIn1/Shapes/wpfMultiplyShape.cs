using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows;
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
public sealed class wpfMultiplyShape : Window, IComponentConnector
{
	[CompilerGenerated]
	private Shape m_A;

	[CompilerGenerated]
	private ShapeRange m_A;

	[CompilerGenerated]
	private RectangleF m_A;

	[CompilerGenerated]
	private bool m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("numRows")]
	private MacNumericUpDown m_A;

	[AccessedThroughProperty("numColumns")]
	[CompilerGenerated]
	private MacNumericUpDown m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("numSpacing")]
	private MacNumericUpDown C;

	private bool m_B;

	private Shape Master
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

	private RectangleF MasterRect
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

	private bool IsMetric
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

	internal virtual MacNumericUpDown numRows
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

	internal virtual MacNumericUpDown numColumns
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

	internal virtual MacNumericUpDown numSpacing
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
			C = value;
		}
	}

	public wpfMultiplyShape(Shape shp)
	{
		//IL_008c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0091: Unknown result type (might be due to invalid IL or missing references)
		//IL_0093: Unknown result type (might be due to invalid IL or missing references)
		//IL_009d: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_0132: Unknown result type (might be due to invalid IL or missing references)
		//IL_013c: Expected O, but got Unknown
		//IL_014b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0155: Expected O, but got Unknown
		//IL_0164: Unknown result type (might be due to invalid IL or missing references)
		//IL_016e: Expected O, but got Unknown
		base.Loaded += wpfMultiplyShape_Loaded;
		base.Closing += wpfMultiplyShape_Closing;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		Master = shp;
		MasterRect = new RectangleF(shp.Left, shp.Top, shp.Width, shp.Height);
		IsMetric = RegionInfo.CurrentRegion.IsMetric;
		SpinnerProperties spinnerProps = Regional.GetSpinnerProps((bool?)IsMetric);
		MacNumericUpDown obj = numSpacing;
		obj.CustomUnit = spinnerProps.CustomUnits;
		obj.NumberDecimalDigits = spinnerProps.Decimals;
		obj.SmallChange = spinnerProps.SmallChange;
		obj.LargeChange = spinnerProps.LargeChange;
		obj.Value = PB.Settings.LastGapSize;
		_ = null;
		numRows.Value = 3.0;
		numColumns.Value = 3.0;
		B();
		numRows.ValueChanged += new MacRangeBaseValueChangedHandler(RowColumnsChanged);
		numColumns.ValueChanged += new MacRangeBaseValueChangedHandler(RowColumnsChanged);
		numSpacing.ValueChanged += new MacRangeBaseValueChangedHandler(SpacingChanged);
	}

	private void wpfMultiplyShape_Loaded(object sender, RoutedEventArgs e)
	{
		Dialogs.GetHwndSource(this).AddHook(A);
		A();
		((UIElement)(object)numRows).Focus();
	}

	private void wpfMultiplyShape_Closing(object sender, CancelEventArgs e)
	{
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		//IL_003f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0049: Expected O, but got Unknown
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0064: Expected O, but got Unknown
		Dialogs.GetHwndSource(this).RemoveHook(A);
		numRows.ValueChanged -= new MacRangeBaseValueChangedHandler(RowColumnsChanged);
		numColumns.ValueChanged -= new MacRangeBaseValueChangedHandler(RowColumnsChanged);
		MacNumericUpDown val = numSpacing;
		val.ValueChanged -= new MacRangeBaseValueChangedHandler(SpacingChanged);
		if (val.Value.HasValue)
		{
			PB.Settings.LastGapSize = (float)val.Value.Value;
		}
		val = null;
		try
		{
			Shapes.Select();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		Master = null;
		MasterRect = default(RectangleF);
		Shapes = null;
	}

	private void OnCloseExecuted(object sender, ExecutedRoutedEventArgs e)
	{
		Close();
	}

	private void A()
	{
		double C = default(double);
		double D = default(double);
		Dialogs.A(Master, Dialogs.DialogPosition.Left, ref C, ref D);
		base.Top = D;
		base.Left = C - base.Width;
	}

	private IntPtr A(IntPtr A, int B, IntPtr C, IntPtr D, ref bool E)
	{
		if (B == Dialogs.A)
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

	private void RowColumnsChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		B();
	}

	private void SpacingChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		B();
	}

	private void B()
	{
		checked
		{
			try
			{
				for (int i = Shapes.Count; i >= 1; i += -1)
				{
					if (Shapes[i] == Master)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Shapes[i].Delete();
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_006a;
					}
					continue;
					end_IL_006a:
					break;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			try
			{
				float num;
				if (IsMetric)
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
					num = clsPublish.CentimetersToPoints((float)numSpacing.Value.Value);
				}
				else
				{
					num = clsPublish.InchesToPoints((float)numSpacing.Value.Value);
				}
				float num2 = num;
				float num3 = MasterRect.Width + num2;
				float num4 = MasterRect.Height + num2;
				List<int> list = new List<int>();
				list.Add(Helpers.A(Master));
				int num5 = (int)Math.Round(numRows.Value.Value);
				for (int j = 1; j <= num5; j++)
				{
					int num6 = (int)Math.Round(numColumns.Value.Value);
					for (int k = 1; k <= num6; k++)
					{
						if (j <= 1)
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
							if (k <= 1)
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
						}
						Shape shape = Master.Duplicate()[1];
						shape.Top = MasterRect.Top + num4 * (float)(j - 1);
						shape.Left = MasterRect.Left + num3 * (float)(k - 1);
						list.Add(Helpers.A(shape));
						shape = null;
					}
				}
				Shapes = NG.A.Application.ActiveWindow.Selection.SlideRange[1].Shapes.Range(list.ToArray());
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			finally
			{
				Shape shape = null;
				List<int> list = null;
			}
		}
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_B)
		{
			this.m_B = true;
			Uri resourceLocator = new Uri(AH.A(91715), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		//IL_003f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0049: Expected O, but got Unknown
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0064: Expected O, but got Unknown
		//IL_0075: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Expected O, but got Unknown
		if (connectionId == 1)
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
					((CommandBinding)target).Executed += OnCloseExecuted;
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
					numRows = (MacNumericUpDown)target;
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
					numColumns = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					numSpacing = (MacNumericUpDown)target;
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
}
