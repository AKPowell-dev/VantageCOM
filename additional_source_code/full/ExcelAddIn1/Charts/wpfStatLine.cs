using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Windows.Media;
using A;
using Foo.Controls;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

[DesignerGenerated]
public sealed class wpfStatLine : Window, IComponentConnector
{
	[AccessedThroughProperty("btnColor")]
	[CompilerGenerated]
	private Button A;

	[CompilerGenerated]
	[AccessedThroughProperty("numWeight")]
	private MacNumericUpDown A;

	[AccessedThroughProperty("cbxStyle")]
	[CompilerGenerated]
	private ComboBox A;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxLabel")]
	private ComboBox B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkColor")]
	private CheckBox A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkBold")]
	private CheckBox B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private Button B;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private Button C;

	private bool A;

	internal virtual Button btnColor
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnColor_Click;
			Button button = this.A;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.A = value;
			button = this.A;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual MacNumericUpDown numWeight
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	internal virtual ComboBox cbxStyle
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	internal virtual ComboBox cbxLabel
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
			SelectionChangedEventHandler value2 = cbxLabel_SelectionChanged;
			ComboBox comboBox = this.B;
			if (comboBox != null)
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
				comboBox.SelectionChanged -= value2;
			}
			this.B = value;
			comboBox = this.B;
			if (comboBox != null)
			{
				comboBox.SelectionChanged += value2;
			}
		}
	}

	internal virtual CheckBox chkColor
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	internal virtual CheckBox chkBold
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

	internal virtual Button btnOk
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
			RoutedEventHandler value2 = btnOk_Click;
			Button button = B;
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
			B = value;
			button = B;
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
			return C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	public wpfStatLine()
	{
		base.Loaded += wpfStatLine_Loaded;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void wpfStatLine_Loaded(object sender, RoutedEventArgs e)
	{
		int selectedIndex = cbxStyle.SelectedIndex;
		cbxStyle.SelectedIndex = -1;
		cbxStyle.SelectedIndex = selectedIndex;
	}

	private void btnColor_Click(object sender, RoutedEventArgs e)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		wpfPalette val = new wpfPalette(false, (ColorRole)0);
		((Window)(object)val).Owner = Window.GetWindow(this);
		((Window)(object)val).Title = VH.A(55164);
		((Window)(object)val).ShowDialog();
		if (((Window)(object)val).DialogResult.HasValue)
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
			if (((Window)(object)val).DialogResult.Value)
			{
				if (val.SelectedColor != Colors.Transparent)
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
					btnColor.Foreground = new SolidColorBrush(val.SelectedColor);
				}
				else
				{
					Forms.WarningMessage(VH.A(55189));
				}
			}
		}
		val = null;
	}

	private void cbxLabel_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (cbxLabel.SelectedIndex == 0)
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
					chkColor.IsEnabled = false;
					chkBold.IsEnabled = false;
					chkColor.IsChecked = false;
					chkBold.IsChecked = false;
					return;
				}
			}
		}
		chkColor.IsEnabled = true;
		chkBold.IsEnabled = true;
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = true;
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (A)
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
			A = true;
			Uri resourceLocator = new Uri(VH.A(69748), UriKind.Relative);
			Application.LoadComponent(this, resourceLocator);
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
		//IL_002a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0034: Expected O, but got Unknown
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
					btnColor = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			numWeight = (MacNumericUpDown)target;
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
					cbxStyle = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			cbxLabel = (ComboBox)target;
			return;
		}
		if (connectionId == 5)
		{
			chkColor = (CheckBox)target;
			return;
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkBold = (CheckBox)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 7:
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				btnOk = (Button)target;
				return;
			}
		case 8:
			btnCancel = (Button)target;
			break;
		default:
			A = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
