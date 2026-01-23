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
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Utilities;

namespace PowerPointAddIn1.Shapes;

[DesignerGenerated]
public sealed class wpfSplitShape : Window, IComponentConnector
{
	[CompilerGenerated]
	private Microsoft.Office.Interop.PowerPoint.Shape m_A;

	[CompilerGenerated]
	private Microsoft.Office.Interop.PowerPoint.ShapeRange m_A;

	[CompilerGenerated]
	private RectangleF m_A;

	[CompilerGenerated]
	private bool m_A;

	[AccessedThroughProperty("numRows")]
	[CompilerGenerated]
	private MacNumericUpDown m_A;

	[AccessedThroughProperty("numColumns")]
	[CompilerGenerated]
	private MacNumericUpDown m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("numSpacing")]
	private MacNumericUpDown m_C;

	private bool m_B;

	private Microsoft.Office.Interop.PowerPoint.Shape Master
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

	private Microsoft.Office.Interop.PowerPoint.ShapeRange Shapes
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
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	public wpfSplitShape(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		//IL_008c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0091: Unknown result type (might be due to invalid IL or missing references)
		//IL_0093: Unknown result type (might be due to invalid IL or missing references)
		//IL_009d: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_0142: Unknown result type (might be due to invalid IL or missing references)
		//IL_014c: Expected O, but got Unknown
		//IL_015b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0165: Expected O, but got Unknown
		//IL_0172: Unknown result type (might be due to invalid IL or missing references)
		//IL_017c: Expected O, but got Unknown
		base.Loaded += wpfSplitShape_Loaded;
		base.Closing += wpfSplitShape_Closing;
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
		Master.Visible = MsoTriState.msoFalse;
		B();
		C();
		numRows.ValueChanged += new MacRangeBaseValueChangedHandler(RowColumnsChanged);
		numColumns.ValueChanged += new MacRangeBaseValueChangedHandler(RowColumnsChanged);
		numSpacing.ValueChanged += new MacRangeBaseValueChangedHandler(SpacingChanged);
	}

	private void wpfSplitShape_Loaded(object sender, RoutedEventArgs e)
	{
		Dialogs.GetHwndSource(this).AddHook(A);
		A();
		((UIElement)(object)numRows).Focus();
	}

	private void wpfSplitShape_Closing(object sender, CancelEventArgs e)
	{
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_0030: Expected O, but got Unknown
		//IL_003d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Expected O, but got Unknown
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		//IL_0060: Expected O, but got Unknown
		Dialogs.GetHwndSource(this).RemoveHook(A);
		numRows.ValueChanged -= new MacRangeBaseValueChangedHandler(RowColumnsChanged);
		numColumns.ValueChanged -= new MacRangeBaseValueChangedHandler(RowColumnsChanged);
		MacNumericUpDown val = numSpacing;
		val.ValueChanged -= new MacRangeBaseValueChangedHandler(SpacingChanged);
		if (val.Value.HasValue)
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
			PB.Settings.LastGapSize = (float)val.Value.Value;
		}
		val = null;
		try
		{
			Master.Delete();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			Shapes.Select();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
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
		C();
	}

	private void SpacingChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		C();
	}

	private void B()
	{
		//IL_008a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0094: Expected O, but got Unknown
		//IL_00ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c4: Expected O, but got Unknown
		checked
		{
			int num = (int)Math.Round(numRows.Value.Value);
			int num2 = (int)Math.Round(numColumns.Value.Value);
			float val = MasterRect.Width / (float)(num2 - 1);
			float val2 = MasterRect.Height / (float)(num - 1);
			float num3 = Math.Min(val, val2);
			MacNumericUpDown obj = numSpacing;
			obj.ValueChanged -= new MacRangeBaseValueChangedHandler(SpacingChanged);
			obj.Maximum = ((!IsMetric) ? clsPublish.PointsToInches(num3) : clsPublish.PointsToCentimeters(num3));
			obj.ValueChanged += new MacRangeBaseValueChangedHandler(SpacingChanged);
			_ = null;
		}
	}

	private void C()
	{
		try
		{
			Shapes.Delete();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		checked
		{
			try
			{
				int num = (int)Math.Round(numRows.Value.Value);
				int num2 = (int)Math.Round(numColumns.Value.Value);
				float num3;
				if (IsMetric)
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
					num3 = clsPublish.CentimetersToPoints((float)numSpacing.Value.Value);
				}
				else
				{
					num3 = clsPublish.InchesToPoints((float)numSpacing.Value.Value);
				}
				float num4 = num3;
				float num5 = (MasterRect.Height - num4 * (float)(num - 1)) / (float)num;
				float num6 = (MasterRect.Width - num4 * (float)(num2 - 1)) / (float)num2;
				float num7 = num5 + num4;
				float num8 = num6 + num4;
				List<int> list = new List<int>();
				int num9 = num - 1;
				for (int i = 0; i <= num9; i++)
				{
					int num10 = num2 - 1;
					for (int j = 0; j <= num10; j++)
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape = Master.Duplicate()[1];
						Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
						shape2.Height = num5;
						shape2.Width = num6;
						shape2.Top = MasterRect.Top + num7 * (float)i;
						shape2.Left = MasterRect.Left + num8 * (float)j;
						Master.ZOrder(MsoZOrderCmd.msoBringToFront);
						list.Add(Helpers.A(shape));
						_ = null;
					}
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					Shapes = NG.A.Application.ActiveWindow.Selection.SlideRange[1].Shapes.Range(list.ToArray());
					Shapes.Visible = MsoTriState.msoTrue;
					return;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			finally
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = null;
				List<int> list = null;
			}
		}
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (this.m_B)
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
			this.m_B = true;
			Uri resourceLocator = new Uri(AH.A(75735), UriKind.Relative);
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
		//IL_0022: Unknown result type (might be due to invalid IL or missing references)
		//IL_002c: Expected O, but got Unknown
		//IL_0057: Unknown result type (might be due to invalid IL or missing references)
		//IL_0061: Expected O, but got Unknown
		//IL_0046: Unknown result type (might be due to invalid IL or missing references)
		//IL_0050: Expected O, but got Unknown
		switch (connectionId)
		{
		case 1:
			((CommandBinding)target).Executed += OnCloseExecuted;
			break;
		case 2:
			numRows = (MacNumericUpDown)target;
			break;
		case 3:
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
				numColumns = (MacNumericUpDown)target;
				return;
			}
		case 4:
			numSpacing = (MacNumericUpDown)target;
			break;
		default:
			this.m_B = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
