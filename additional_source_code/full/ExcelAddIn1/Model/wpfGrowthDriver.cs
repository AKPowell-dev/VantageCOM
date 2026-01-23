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
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Model;

[DesignerGenerated]
public sealed class wpfGrowthDriver : System.Windows.Window, IComponentConnector
{
	[AccessedThroughProperty("cbxColumn")]
	[CompilerGenerated]
	private ComboBox m_A;

	[AccessedThroughProperty("chkGrowth")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lblGrowth")]
	private TextBlock m_A;

	[AccessedThroughProperty("txtLabel")]
	[CompilerGenerated]
	private TextBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkIndent")]
	private System.Windows.Controls.CheckBox B;

	[CompilerGenerated]
	[AccessedThroughProperty("lblFormat")]
	private TextBlock B;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxNumFormat")]
	private ComboBox B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkLookup")]
	private System.Windows.Controls.CheckBox C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private Button B;

	private bool m_A;

	internal virtual ComboBox cbxColumn
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

	internal virtual System.Windows.Controls.CheckBox chkGrowth
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
			RoutedEventHandler value2 = chkGrowth_CheckedChanged;
			RoutedEventHandler value3 = chkGrowth_CheckedChanged;
			System.Windows.Controls.CheckBox checkBox = this.m_A;
			if (checkBox != null)
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
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.m_A = value;
			checkBox = this.m_A;
			if (checkBox == null)
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
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual TextBlock lblGrowth
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

	internal virtual TextBox txtLabel
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

	internal virtual System.Windows.Controls.CheckBox chkIndent
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	internal virtual TextBlock lblFormat
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	internal virtual ComboBox cbxNumFormat
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkLookup
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

	internal virtual Button btnOk
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
			RoutedEventHandler value2 = btnOk_Click;
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
				switch (2)
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
			return B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public wpfGrowthDriver()
	{
		base.Loaded += wpfGrowthDriver_Loaded;
		base.Closing += wpfGrowthDriver_Closing;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void wpfGrowthDriver_Loaded(object sender, RoutedEventArgs e)
	{
		cbxColumn.SelectionChanged += cbxColumn_SelectionChanged;
	}

	private void wpfGrowthDriver_Closing(object sender, CancelEventArgs e)
	{
		cbxColumn.SelectionChanged -= cbxColumn_SelectionChanged;
	}

	private void cbxColumn_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		Range range = (Range)MH.A.Application.Selection;
		string text = Strings.Replace(cbxColumn.SelectedItem.ToString(), VH.A(91534), "");
		((_Worksheet)range.Worksheet).get_Range((object)(text + Conversions.ToString(range.Row)), RuntimeHelpers.GetObjectValue(Missing.Value)).Activate();
		range = null;
	}

	private void chkGrowth_CheckedChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked = chkGrowth.IsChecked;
		bool? flag;
		if (!isChecked.HasValue)
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
			flag = isChecked;
		}
		else
		{
			flag = isChecked != true;
		}
		isChecked = flag;
		A(isChecked.Value);
	}

	private void A(bool A)
	{
		lblGrowth.IsEnabled = A;
		lblFormat.IsEnabled = A;
		txtLabel.IsEnabled = A;
		chkIndent.IsEnabled = A;
		cbxNumFormat.IsEnabled = A;
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = true;
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
			switch (2)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_A = true;
			Uri resourceLocator = new Uri(VH.A(91549), UriKind.Relative);
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
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
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
					cbxColumn = (ComboBox)target;
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
					chkGrowth = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			lblGrowth = (TextBlock)target;
			return;
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					txtLabel = (TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkIndent = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					lblFormat = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					cbxNumFormat = (ComboBox)target;
					return;
				}
			}
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
					chkLookup = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 9:
			btnOk = (Button)target;
			break;
		case 10:
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				btnCancel = (Button)target;
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
