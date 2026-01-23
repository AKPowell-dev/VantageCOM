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
public sealed class wpfTargetBand : Window, IComponentConnector
{
	[CompilerGenerated]
	[AccessedThroughProperty("btnColor")]
	private Button m_A;

	[AccessedThroughProperty("numTransparency")]
	[CompilerGenerated]
	private MacNumericUpDown m_A;

	[AccessedThroughProperty("txtUpper")]
	[CompilerGenerated]
	private TextBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("txtLower")]
	private TextBox B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private Button B;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private Button C;

	private bool m_A;

	internal virtual Button btnColor
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
			RoutedEventHandler value2 = btnColor_Click;
			Button button = this.m_A;
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
			this.m_A = value;
			button = this.m_A;
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

	internal virtual MacNumericUpDown numTransparency
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

	internal virtual TextBox txtUpper
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
			RoutedEventHandler value2 = txtLower_GotFocus;
			TextBox textBox = this.m_A;
			if (textBox != null)
			{
				textBox.GotFocus -= value2;
			}
			this.m_A = value;
			textBox = this.m_A;
			if (textBox == null)
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
				textBox.GotFocus += value2;
				return;
			}
		}
	}

	internal virtual TextBox txtLower
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
			RoutedEventHandler value2 = txtLower_GotFocus;
			TextBox textBox = this.B;
			if (textBox != null)
			{
				textBox.GotFocus -= value2;
			}
			this.B = value;
			textBox = this.B;
			if (textBox == null)
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
				textBox.GotFocus += value2;
				return;
			}
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
			B = value;
			button = B;
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
			return C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	public wpfTargetBand()
	{
		base.Loaded += wpfTargetBand_Loaded;
		base.Closing += wpfTargetBand_Closing;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void wpfTargetBand_Loaded(object sender, RoutedEventArgs e)
	{
		txtUpper.SelectAll();
		txtUpper.Focus();
	}

	private void wpfTargetBand_Closing(object sender, CancelEventArgs e)
	{
		if (base.DialogResult != true)
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
			if (!float.TryParse(txtUpper.Text, out var result))
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						A(VH.A(68093));
						e.Cancel = true;
						return;
					}
				}
			}
			if (!float.TryParse(txtLower.Text, out var result2))
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						A(VH.A(68150));
						e.Cancel = true;
						return;
					}
				}
			}
			if (result <= result2)
			{
				A(VH.A(68207));
				e.Cancel = true;
				return;
			}
			if (!(result2 < 0f))
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
				A(VH.A(68296));
				e.Cancel = true;
				return;
			}
		}
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
			if (((Window)(object)val).DialogResult.Value)
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
				if (val.SelectedColor != Colors.Transparent)
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

	private void txtLower_GotFocus(object sender, RoutedEventArgs e)
	{
		((TextBox)sender).SelectAll();
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = true;
	}

	private void A(string A)
	{
		Forms.WarningMessage(Window.GetWindow(this), A);
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
			Uri resourceLocator = new Uri(VH.A(68373), UriKind.Relative);
			Application.LoadComponent(this, resourceLocator);
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
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003e: Expected O, but got Unknown
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
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					numTransparency = (MacNumericUpDown)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 3:
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				txtUpper = (TextBox)target;
				return;
			}
		case 4:
			txtLower = (TextBox)target;
			break;
		case 5:
			btnOk = (Button)target;
			break;
		case 6:
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
