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

namespace ExcelAddIn1.Audit.Visualizations;

[DesignerGenerated]
public sealed class wpfMagnitudeMap : Window, IComponentConnector
{
	[CompilerGenerated]
	[AccessedThroughProperty("radAll")]
	private RadioButton A;

	[AccessedThroughProperty("radRows")]
	[CompilerGenerated]
	private RadioButton B;

	[AccessedThroughProperty("radColumns")]
	[CompilerGenerated]
	private RadioButton C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private Button A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private Button B;

	private bool A;

	internal virtual RadioButton radAll
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

	internal virtual RadioButton radRows
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

	internal virtual RadioButton radColumns
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
			RoutedEventHandler value2 = CheckOnFocus;
			RadioButton radioButton = C;
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
			C = value;
			radioButton = C;
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

	public wpfMagnitudeMap()
	{
		base.Loaded += wpfPasteTrans_Loaded;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void wpfPasteTrans_Loaded(object sender, RoutedEventArgs e)
	{
		radAll.Focus();
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
			Uri resourceLocator = new Uri(VH.A(51603), UriKind.Relative);
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
	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		switch (connectionId)
		{
		case 1:
			radAll = (RadioButton)target;
			break;
		case 2:
			radRows = (RadioButton)target;
			break;
		case 3:
			radColumns = (RadioButton)target;
			break;
		case 4:
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
				btnOk = (Button)target;
				return;
			}
		case 5:
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
