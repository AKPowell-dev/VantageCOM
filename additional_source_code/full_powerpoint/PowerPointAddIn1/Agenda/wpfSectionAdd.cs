using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using A;
using MacabacusMacros.UI;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Agenda;

[DesignerGenerated]
public sealed class wpfSectionAdd : Window, IComponentConnector
{
	[CompilerGenerated]
	[AccessedThroughProperty("txtTopic")]
	private TextBox A;

	[AccessedThroughProperty("radSection")]
	[CompilerGenerated]
	private RadioButton A;

	[AccessedThroughProperty("radSubsection")]
	[CompilerGenerated]
	private RadioButton B;

	[AccessedThroughProperty("radAppendix")]
	[CompilerGenerated]
	private RadioButton C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private Button A;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private Button B;

	private bool A;

	internal virtual TextBox txtTopic
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

	internal virtual RadioButton radSection
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

	internal virtual RadioButton radSubsection
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

	internal virtual RadioButton radAppendix
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
			return this.A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnOk_Click;
			Button button = this.A;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.A = value;
			button = this.A;
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
			return B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public wpfSectionAdd()
	{
		base.Loaded += wpfSectionAdd_Loadws;
		base.Closing += wpfFlysheet_Closing;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void wpfSectionAdd_Loadws(object sender, RoutedEventArgs e)
	{
		TextBox textBox = txtTopic;
		textBox.Text = AH.A(7090);
		textBox.Focus();
		textBox.SelectAll();
		_ = null;
		txtTopic.TextChanged += txtTopic_TextChanged;
		radSection.Checked += radSection_CheckedChanged;
		radSection.Unchecked += radSection_CheckedChanged;
	}

	private void txtTopic_TextChanged(object sender, TextChangedEventArgs e)
	{
		radSection.Checked -= radSection_CheckedChanged;
		radSection.Unchecked -= radSection_CheckedChanged;
		if (txtTopic.Text.StartsWith(Constants.SUBSECTION_PREFIX))
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
			radSubsection.IsChecked = true;
		}
		else
		{
			radSection.IsChecked = true;
		}
		radSection.Checked += radSection_CheckedChanged;
		radSection.Unchecked += radSection_CheckedChanged;
	}

	private void radSection_CheckedChanged(object sender, RoutedEventArgs e)
	{
		TextBox textBox = txtTopic;
		textBox.TextChanged -= txtTopic_TextChanged;
		if (radSection.IsChecked == true)
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
			if (textBox.Text.StartsWith(Constants.SUBSECTION_PREFIX))
			{
				textBox.Text = Strings.Mid(textBox.Text, 2);
			}
		}
		else
		{
			textBox.Text = Constants.SUBSECTION_PREFIX + textBox.Text;
		}
		textBox.TextChanged += txtTopic_TextChanged;
		textBox = null;
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		if (txtTopic.Text.Length == 0)
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
					Forms.WarningMessage(AH.A(7123));
					return;
				}
			}
		}
		base.DialogResult = true;
	}

	private void wpfFlysheet_Closing(object sender, CancelEventArgs e)
	{
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (A)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A = true;
			Uri resourceLocator = new Uri(AH.A(7162), UriKind.Relative);
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
					txtTopic = (TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			radSection = (RadioButton)target;
			return;
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					radSubsection = (RadioButton)target;
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
					radAppendix = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnOk = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnCancel = (Button)target;
					return;
				}
			}
		}
		A = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
