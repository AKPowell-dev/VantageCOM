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
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

[DesignerGenerated]
public sealed class wpfErrorWrap : Window, IComponentConnector
{
	[AccessedThroughProperty("optAdd")]
	[CompilerGenerated]
	private RadioButton A;

	[CompilerGenerated]
	[AccessedThroughProperty("optRemove")]
	private RadioButton B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private Button A;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private Button B;

	private bool A;

	internal virtual RadioButton optAdd
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
			RoutedEventHandler value2 = CheckOnFocus;
			RadioButton radioButton = this.A;
			if (radioButton != null)
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
				radioButton.GotFocus -= value2;
			}
			this.A = value;
			radioButton = this.A;
			if (radioButton != null)
			{
				radioButton.GotFocus += value2;
			}
		}
	}

	internal virtual RadioButton optRemove
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
			RoutedEventHandler value2 = CheckOnFocus;
			RadioButton radioButton = this.B;
			if (radioButton != null)
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
				radioButton.GotFocus -= value2;
			}
			this.B = value;
			radioButton = this.B;
			if (radioButton != null)
			{
				radioButton.GotFocus += value2;
			}
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

	public wpfErrorWrap()
	{
		base.Loaded += wpfErrorWrap_Loaded;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void wpfErrorWrap_Loaded(object sender, RoutedEventArgs e)
	{
		if (optAdd.IsChecked == true)
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
					optAdd.Focus();
					return;
				}
			}
		}
		optRemove.Focus();
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = true;
	}

	private void CheckOnFocus(object sender, RoutedEventArgs e)
	{
		((RadioButton)sender).IsChecked = true;
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
			switch (6)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A = true;
			Uri resourceLocator = new Uri(VH.A(157803), UriKind.Relative);
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
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					optAdd = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					optRemove = (RadioButton)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 3:
			btnOk = (Button)target;
			break;
		case 4:
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				btnCancel = (Button)target;
				return;
			}
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
