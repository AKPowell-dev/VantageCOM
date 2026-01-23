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
public sealed class wpfStackedColumn : System.Windows.Window, IComponentConnector
{
	public Range Range;

	private Worksheet m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("txtAddress")]
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

	[AccessedThroughProperty("numChartHeight")]
	[CompilerGenerated]
	private MacNumericUpDown m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("numChartWidth")]
	private MacNumericUpDown m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkShowAmount")]
	private System.Windows.Controls.CheckBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkShowPercent")]
	private System.Windows.Controls.CheckBox m_C;

	[AccessedThroughProperty("chkTotalsShow")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_D;

	[AccessedThroughProperty("chkTotalsBold")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox E;

	[CompilerGenerated]
	[AccessedThroughProperty("numGapWidth")]
	private MacNumericUpDown m_C;

	[AccessedThroughProperty("btnOk")]
	[CompilerGenerated]
	private Button m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private Button m_D;

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
			this.m_A = value;
			button = this.m_A;
			if (button == null)
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
				switch (3)
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

	internal virtual System.Windows.Controls.CheckBox chkShowAmount
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

	internal virtual System.Windows.Controls.CheckBox chkShowPercent
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

	internal virtual System.Windows.Controls.CheckBox chkTotalsShow
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
			RoutedEventHandler value2 = chkTotalsShow_CheckedChanged;
			RoutedEventHandler value3 = chkTotalsShow_CheckedChanged;
			System.Windows.Controls.CheckBox checkBox = this.m_D;
			if (checkBox != null)
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
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.m_D = value;
			checkBox = this.m_D;
			if (checkBox == null)
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
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkTotalsBold
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

	internal virtual Button btnOk
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
			RoutedEventHandler value2 = btnOk_Click;
			Button button = this.m_C;
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
			this.m_C = value;
			button = this.m_C;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual Button btnCancel
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

	public wpfStackedColumn()
	{
		base.Loaded += wpfStackedColumn_Loaded;
		base.Closing += wpfStackedColumn_Closing;
		Range = null;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void wpfStackedColumn_Loaded(object sender, RoutedEventArgs e)
	{
		btnOk.Focus();
		QuickCharts2.ChartDialogLoad(cbxChartSize, numChartHeight, numChartWidth);
		A();
		C();
	}

	private void wpfStackedColumn_Closing(object sender, CancelEventArgs e)
	{
		bool? dialogResult = base.DialogResult;
		bool? obj;
		if (!dialogResult.HasValue)
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
				switch (4)
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

	private void btnRangeEdit_Click(object sender, EventArgs e)
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
					switch (7)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (!Operators.ConditionalCompareObjectLess(range.Columns.CountLarge, 2, TextCompare: false))
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
						if (!Operators.ConditionalCompareObjectLess(range.Rows.CountLarge, 2, TextCompare: false))
						{
							txtAddress.Text = range.get_Address((object)true, (object)true, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							Range = range;
							break;
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
					}
					Forms.WarningMessage(VH.A(74553));
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
		Range = QuickCharts2.SampleChartData(txtAddress, StackedColumn.Example);
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = true;
	}

	private void chkTotalsShow_CheckedChanged(object sender, RoutedEventArgs e)
	{
		chkTotalsBold.IsEnabled = chkTotalsShow.IsChecked.Value;
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
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_A = true;
			Uri resourceLocator = new Uri(VH.A(74666), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
			return;
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		//IL_008c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0096: Expected O, but got Unknown
		//IL_00a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b1: Expected O, but got Unknown
		//IL_0114: Unknown result type (might be due to invalid IL or missing references)
		//IL_011e: Expected O, but got Unknown
		if (connectionId == 1)
		{
			while (true)
			{
				switch (6)
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
			btnRangeEdit = (Button)target;
			return;
		}
		if (connectionId == 3)
		{
			btnExample = (Button)target;
			return;
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (1)
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
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					cbxChartSize = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (7)
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
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					numChartWidth = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			chkShowAmount = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 9)
		{
			chkShowPercent = (System.Windows.Controls.CheckBox)target;
			return;
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
					chkTotalsShow = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkTotalsBold = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			numGapWidth = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnOk = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnCancel = (Button)target;
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
