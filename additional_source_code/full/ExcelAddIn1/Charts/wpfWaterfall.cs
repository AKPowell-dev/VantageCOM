using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using A;
using Foo.Controls;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

[DesignerGenerated]
public sealed class wpfWaterfall : System.Windows.Window, IComponentConnector
{
	public Range Range;

	private Worksheet m_A;

	[AccessedThroughProperty("txtAddress")]
	[CompilerGenerated]
	private TextBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnRangeEdit")]
	private Button m_A;

	[AccessedThroughProperty("btnExample")]
	[CompilerGenerated]
	private Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPreserveFormulas")]
	private System.Windows.Controls.CheckBox m_A;

	[AccessedThroughProperty("cbxChartSize")]
	[CompilerGenerated]
	private ComboBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("numChartHeight")]
	private MacNumericUpDown m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("numChartWidth")]
	private MacNumericUpDown m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("radLabelCenter")]
	private RadioButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("radLabelAboveBelow")]
	private RadioButton m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("radLabelAbove")]
	private RadioButton m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("radLabelNone")]
	private RadioButton m_D;

	[AccessedThroughProperty("btnColorBase")]
	[CompilerGenerated]
	private Button m_C;

	[AccessedThroughProperty("btnColorIncrease")]
	[CompilerGenerated]
	private Button m_D;

	[AccessedThroughProperty("btnColorDecrease")]
	[CompilerGenerated]
	private Button E;

	[CompilerGenerated]
	[AccessedThroughProperty("numGapWidth")]
	private MacNumericUpDown m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("chkErrorBars")]
	private System.Windows.Controls.CheckBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private Button F;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private Button G;

	private bool m_A;

	internal virtual TextBox txtAddress
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

	internal virtual Button btnRangeEdit
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
			RoutedEventHandler value2 = btnRangeEdit_Click;
			Button button = this.m_A;
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

	internal virtual Button btnExample
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
			RoutedEventHandler value2 = btnExample_Click;
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

	internal virtual System.Windows.Controls.CheckBox chkPreserveFormulas
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

	internal virtual ComboBox cbxChartSize
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

	internal virtual MacNumericUpDown numChartHeight
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

	internal virtual MacNumericUpDown numChartWidth
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

	internal virtual RadioButton radLabelCenter
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

	internal virtual RadioButton radLabelAboveBelow
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

	internal virtual RadioButton radLabelAbove
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

	internal virtual RadioButton radLabelNone
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

	internal virtual Button btnColorBase
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

	internal virtual Button btnColorIncrease
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

	internal virtual Button btnColorDecrease
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			E = value;
		}
	}

	internal virtual MacNumericUpDown numGapWidth
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

	internal virtual System.Windows.Controls.CheckBox chkErrorBars
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

	internal virtual Button btnOk
	{
		[CompilerGenerated]
		get
		{
			return F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnOk_Click;
			Button button = F;
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
			F = value;
			button = F;
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

	internal virtual Button btnCancel
	{
		[CompilerGenerated]
		get
		{
			return G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			G = value;
		}
	}

	public wpfWaterfall()
	{
		base.Loaded += wpfWaterfall_Loaded;
		base.Closing += wpfWaterfall_Closing;
		Range = null;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void wpfWaterfall_Loaded(object sender, RoutedEventArgs e)
	{
		btnOk.Focus();
		QuickCharts2.ChartDialogLoad(cbxChartSize, numChartHeight, numChartWidth);
		A();
		C();
	}

	private void wpfWaterfall_Closing(object sender, CancelEventArgs e)
	{
		bool? dialogResult = base.DialogResult;
		bool? obj;
		if (!dialogResult.HasValue)
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
			obj = null;
		}
		else
		{
			obj = dialogResult != true;
		}
		dialogResult = obj;
		if (dialogResult == true)
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
			QuickCharts2.DeleteSampleWorksheet(this.m_A);
		}
		this.m_A = null;
		B();
		D();
	}

	private void btnRangeEdit_Click(object sender, RoutedEventArgs e)
	{
		string text = "";
		try
		{
			text = Range.get_Address((object)true, (object)true, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			Range range = (Range)MH.A.Application.InputBox(VH.A(62623), VH.A(40448), text, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), 8);
			if (range != null)
			{
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
					if (Operators.ConditionalCompareObjectNotEqual(range.Columns.CountLarge, 2, TextCompare: false))
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							Forms.WarningMessage(VH.A(74830));
							break;
						}
					}
					else
					{
						txtAddress.Text = range.get_Address((object)true, (object)true, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						Range = range;
					}
					break;
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		Focus();
		btnOk.Focus();
	}

	private void btnExample_Click(object sender, RoutedEventArgs e)
	{
		Range = QuickCharts2.SampleChartData(txtAddress, Waterfall.Example);
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = true;
	}

	private void StandardSizeSelectedIndexChanged(object sender, SelectionChangedEventArgs e)
	{
		D();
		QuickCharts2.StandardSizeSelectedIndexChanged(cbxChartSize, numChartHeight, numChartWidth);
		C();
	}

	private void ChartSizeValueChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		B();
		D();
		QuickCharts2.CheckForStandardSize(cbxChartSize, numChartHeight, numChartWidth);
		A();
		C();
	}

	private void A()
	{
		cbxChartSize.SelectionChanged += StandardSizeSelectedIndexChanged;
	}

	private void B()
	{
		cbxChartSize.SelectionChanged -= StandardSizeSelectedIndexChanged;
	}

	private void C()
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Expected O, but got Unknown
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		numChartHeight.ValueChanged += new MacRangeBaseValueChangedHandler(ChartSizeValueChanged);
		numChartWidth.ValueChanged += new MacRangeBaseValueChangedHandler(ChartSizeValueChanged);
	}

	private void D()
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Expected O, but got Unknown
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		numChartHeight.ValueChanged -= new MacRangeBaseValueChangedHandler(ChartSizeValueChanged);
		numChartWidth.ValueChanged -= new MacRangeBaseValueChangedHandler(ChartSizeValueChanged);
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (this.m_A)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_A = true;
			Uri resourceLocator = new Uri(VH.A(74907), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
			return;
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
		//IL_00a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b1: Expected O, but got Unknown
		//IL_0096: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a0: Expected O, but got Unknown
		//IL_015e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0168: Expected O, but got Unknown
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
					txtAddress = (TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnRangeEdit = (Button)target;
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
					btnExample = (Button)target;
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
					chkPreserveFormulas = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			cbxChartSize = (ComboBox)target;
			return;
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					numChartHeight = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			numChartWidth = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					radLabelCenter = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			radLabelAboveBelow = (RadioButton)target;
			return;
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					radLabelAbove = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			radLabelNone = (RadioButton)target;
			return;
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnColorBase = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 13)
		{
			btnColorIncrease = (Button)target;
			return;
		}
		if (connectionId == 14)
		{
			btnColorDecrease = (Button)target;
			return;
		}
		if (connectionId == 15)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					numGapWidth = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkErrorBars = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 17:
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				btnOk = (Button)target;
				return;
			}
		case 18:
			btnCancel = (Button)target;
			break;
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
