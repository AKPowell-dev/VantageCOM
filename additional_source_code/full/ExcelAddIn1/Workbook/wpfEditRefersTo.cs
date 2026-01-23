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

namespace ExcelAddIn1.Workbook;

[DesignerGenerated]
public sealed class wpfEditRefersTo : Window, IComponentConnector
{
	[CompilerGenerated]
	[AccessedThroughProperty("scv")]
	private ScrollViewer A;

	[AccessedThroughProperty("txtRefersTo")]
	[CompilerGenerated]
	private TextBox A;

	[AccessedThroughProperty("btnSave")]
	[CompilerGenerated]
	private Button A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private Button B;

	private bool A;

	internal virtual ScrollViewer scv
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

	internal virtual TextBox txtRefersTo
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

	internal virtual Button btnSave
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
			RoutedEventHandler value2 = btnSave_Click;
			Button button = this.A;
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
			this.A = value;
			button = this.A;
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

	public wpfEditRefersTo()
	{
		base.Closing += wpfEditRefersTo_Closing;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void wpfEditRefersTo_Closing(object sender, CancelEventArgs e)
	{
		if (base.DialogResult != true)
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
			if (txtRefersTo.Text.StartsWith(VH.A(48936)))
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
				Forms.WarningMessage(VH.A(186321));
				e.Cancel = true;
				return;
			}
		}
	}

	private void btnSave_Click(object sender, RoutedEventArgs e)
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
			switch (2)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A = true;
			Uri resourceLocator = new Uri(VH.A(186392), UriKind.Relative);
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
					scv = (ScrollViewer)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 2:
			txtRefersTo = (TextBox)target;
			break;
		case 3:
			btnSave = (Button)target;
			break;
		case 4:
			while (true)
			{
				switch (7)
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
