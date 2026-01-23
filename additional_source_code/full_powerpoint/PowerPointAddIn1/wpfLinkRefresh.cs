using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using A;
using MacabacusMacros.Links;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1;

[DesignerGenerated]
public sealed class wpfLinkRefresh : Window, IComponentConnector
{
	public int ProgressPercent;

	public bool Canceled;

	public List<Slide> SlideLinks;

	public List<Shape> ShapeLinks;

	public List<TextLink> TextLinks;

	public List<Hyperlink> Hyperlinks;

	public bool RefreshAll;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private Button A;

	[AccessedThroughProperty("pbLink")]
	[CompilerGenerated]
	private ProgressBar A;

	[CompilerGenerated]
	[AccessedThroughProperty("tbStatus")]
	private TextBlock A;

	private bool A;

	internal virtual Button btnCancel
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
			MouseEventHandler value2 = btnCancel_MouseEnter;
			MouseEventHandler value3 = btnCancel_MouseLeave;
			RoutedEventHandler value4 = btnCancel_Click;
			Button button = this.A;
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
				button.MouseEnter -= value2;
				button.MouseLeave -= value3;
				button.Click -= value4;
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
				button.MouseEnter += value2;
				button.MouseLeave += value3;
				button.Click += value4;
				return;
			}
		}
	}

	internal virtual ProgressBar pbLink
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

	internal virtual TextBlock tbStatus
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

	public wpfLinkRefresh()
	{
		base.MouseDown += Window_MouseDown;
		ProgressPercent = 0;
		Canceled = false;
		SlideLinks = null;
		ShapeLinks = null;
		TextLinks = null;
		Hyperlinks = null;
		RefreshAll = false;
		InitializeComponent();
	}

	private void btnCancel_MouseEnter(object sender, MouseEventArgs e)
	{
		btnCancel.Opacity = 1.0;
	}

	private void btnCancel_MouseLeave(object sender, MouseEventArgs e)
	{
		btnCancel.Opacity = 0.6;
	}

	private void btnCancel_Click(object sender, RoutedEventArgs e)
	{
		Canceled = true;
	}

	private void Window_MouseDown(object sender, MouseButtonEventArgs e)
	{
		if (e.ChangedButton != MouseButton.Left)
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
			DragMove();
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
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A = true;
			Uri resourceLocator = new Uri(AH.A(139764), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
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
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					btnCancel = (Button)target;
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
					pbLink = (ProgressBar)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					tbStatus = (TextBlock)target;
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
