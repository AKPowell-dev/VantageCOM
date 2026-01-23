using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using A;
using MacabacusMacros.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.View;

[DesignerGenerated]
public sealed class wpfPrintArea : Window, IComponentConnector
{
	public bool SetPrintArea;

	[AccessedThroughProperty("btnSet")]
	[CompilerGenerated]
	private Button A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnAdd")]
	private Button B;

	private bool A;

	internal virtual Button btnSet
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
			RoutedEventHandler value2 = btnSet_Click;
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

	internal virtual Button btnAdd
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
			RoutedEventHandler value2 = btnAdd_Click;
			Button button = B;
			if (button != null)
			{
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

	public wpfPrintArea()
	{
		base.Loaded += wpfPrintArea_Loaded;
		base.KeyDown += wpfPrintArea_KeyDown;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void wpfPrintArea_Loaded(object sender, RoutedEventArgs e)
	{
		btnSet.Focus();
	}

	private void btnSet_Click(object sender, RoutedEventArgs e)
	{
		SetPrintArea = true;
		base.DialogResult = true;
	}

	private void btnAdd_Click(object sender, RoutedEventArgs e)
	{
		SetPrintArea = false;
		base.DialogResult = true;
	}

	private void wpfPrintArea_KeyDown(object sender, KeyEventArgs e)
	{
		if (e.Key != Key.Escape)
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
			base.DialogResult = false;
			return;
		}
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
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A = true;
			Uri resourceLocator = new Uri(VH.A(174963), UriKind.Relative);
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
		switch (connectionId)
		{
		case 1:
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
				btnSet = (Button)target;
				return;
			}
		case 2:
			btnAdd = (Button)target;
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
