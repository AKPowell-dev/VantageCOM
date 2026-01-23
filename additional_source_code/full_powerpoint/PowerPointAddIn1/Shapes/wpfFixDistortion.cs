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

namespace PowerPointAddIn1.Shapes;

[DesignerGenerated]
public sealed class wpfFixDistortion : Window, IComponentConnector
{
	[AccessedThroughProperty("radAdjHeight")]
	[CompilerGenerated]
	private RadioButton A;

	[CompilerGenerated]
	[AccessedThroughProperty("radAdjWidth")]
	private RadioButton B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private Button A;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private Button B;

	private bool A;

	internal virtual RadioButton radAdjHeight
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

	internal virtual RadioButton radAdjWidth
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

	public wpfFixDistortion()
	{
		base.Loaded += wpfFixDistortion_Loaded;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void wpfFixDistortion_Loaded(object sender, RoutedEventArgs e)
	{
		radAdjHeight.Focus();
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = true;
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
			Uri resourceLocator = new Uri(AH.A(75505), UriKind.Relative);
			Application.LoadComponent(this, resourceLocator);
			return;
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
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
					radAdjHeight = (RadioButton)target;
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
					radAdjWidth = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnOk = (Button)target;
					return;
				}
			}
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
