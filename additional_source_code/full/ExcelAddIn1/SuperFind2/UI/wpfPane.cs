using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using A;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.UI;

[DesignerGenerated]
public sealed class wpfPane : UserControl, INotifyPropertyChanged, IComponentConnector
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private wpfSearch m_A;

	private wpfResults m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("conSearch")]
	private ContentControl m_A;

	[AccessedThroughProperty("conResults")]
	[CompilerGenerated]
	private ContentControl m_B;

	private bool m_A;

	public wpfSearch SearchView
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(140511));
		}
	}

	public wpfResults ResultsView
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(140532));
		}
	}

	internal virtual ContentControl conSearch
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

	internal virtual ContentControl conResults
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	public event PropertyChangedEventHandler PropertyChanged
	{
		[CompilerGenerated]
		add
		{
			PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Combine(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.m_A, value2, propertyChangedEventHandler2);
			}
			while ((object)propertyChangedEventHandler != propertyChangedEventHandler2);
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
				return;
			}
		}
		[CompilerGenerated]
		remove
		{
			PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Remove(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.m_A, value2, propertyChangedEventHandler2);
			}
			while ((object)propertyChangedEventHandler != propertyChangedEventHandler2);
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
				return;
			}
		}
	}

	public wpfPane()
	{
		base.Loaded += PaneLoaded;
		this.m_A = null;
		this.m_A = null;
		InitializeComponent();
		Props.Icons = new IconCache();
		SearchView = new wpfSearch(this);
		A();
	}

	private void A(string A)
	{
		this.m_A?.Invoke(this, new PropertyChangedEventArgs(A));
	}

	private void PaneLoaded(object sender, RoutedEventArgs e)
	{
		clsPanes.EliminateTaskPaneFlicker(this);
	}

	internal void A()
	{
		conSearch.Visibility = Visibility.Visible;
		conResults.Visibility = Visibility.Collapsed;
		ResultsView = null;
	}

	internal void B()
	{
		conResults.Visibility = Visibility.Visible;
		conSearch.Visibility = Visibility.Collapsed;
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
			switch (6)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_A = true;
			Uri resourceLocator = new Uri(VH.A(140555), UriKind.Relative);
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
				conSearch = (ContentControl)target;
				return;
			}
		case 2:
			conResults = (ContentControl)target;
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
