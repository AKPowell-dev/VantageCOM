using System;
using System.CodeDom.Compiler;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using A;
using MacabacusMacros;
using MacabacusMacros.Aiwa;
using MacabacusMacros.Aiwa.UI;
using MacabacusMacros.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Aiwa;

[DesignerGenerated]
public sealed class wpfHome : UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private UserControl m_A;

	private UserControl m_B;

	private UserControl C;

	private ObservableCollection<JsonFeature> m_A;

	[AccessedThroughProperty("scroller")]
	[CompilerGenerated]
	private ScrollViewer m_A;

	private bool m_A;

	public UserControl Workspace
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(XC.A(1074));
		}
	}

	public UserControl ErrorView
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(XC.A(1093));
		}
	}

	public UserControl LoadingView
	{
		get
		{
			return C;
		}
		set
		{
			C = value;
			A(XC.A(1112));
		}
	}

	public ObservableCollection<JsonFeature> Workspaces
	{
		get
		{
			return this.m_A;
		}
		set
		{
			if (object.Equals(this.m_A, value))
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
				this.m_A = value;
				A(XC.A(1135));
				return;
			}
		}
	}

	internal virtual ScrollViewer scroller
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
				switch (1)
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
		}
	}

	public wpfHome()
	{
		base.Loaded += wpfAiwa_Loaded;
		base.Unloaded += wpfAiwa_Unloaded;
		this.m_A = null;
		this.m_B = null;
		C = null;
		InitializeComponent();
	}

	private void A(string A)
	{
		PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
		if (propertyChangedEventHandler == null)
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
			propertyChangedEventHandler(this, new PropertyChangedEventArgs(A));
			return;
		}
	}

	private void wpfAiwa_Loaded(object sender, RoutedEventArgs e)
	{
	}

	private void wpfAiwa_Unloaded(object sender, RoutedEventArgs e)
	{
		Workspace = null;
		ErrorView = null;
		LoadingView = null;
	}

	private void PaneSizeChanged(object sender, SizeChangedEventArgs e)
	{
		Panes.PaneSizeChanged(scroller, e);
	}

	public void ShowPane()
	{
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Expected O, but got Unknown
		base.SizeChanged += PaneSizeChanged;
		string empty = string.Empty;
		LoadingView = (UserControl)new wpfLoading();
		Workspaces = new ObservableCollection<JsonFeature>();
		try
		{
			Workspaces = new ObservableCollection<JsonFeature>(Api.GetAiwaFeatures(ref empty));
		}
		finally
		{
			LoadingView = null;
		}
		if (!modFunctionsStr.IsBlank(empty))
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					B(empty);
					return;
				}
			}
		}
		if (Workspaces.Count != 0)
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
			B(Pane.NOT_CONFIGURED_MSG);
			return;
		}
	}

	public void HidePane()
	{
		Workspace = null;
		base.SizeChanged -= PaneSizeChanged;
	}

	private void LoadWorkspace(object sender, RoutedEventArgs e)
	{
		if (sender is Button button)
		{
			object tag = button.Tag;
			JsonFeature val = (JsonFeature)((tag is JsonFeature) ? tag : null);
			if (val != null)
			{
				Workspace = new wpfWorkspace(this, val);
			}
		}
	}

	internal void A()
	{
		Workspace = null;
	}

	private void B(string A)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		ErrorView = (UserControl)new wpfError(A, (Action<object, RoutedEventArgs>)DismissErrorAndClosePane);
	}

	private void DismissErrorAndClosePane(object sender, RoutedEventArgs e)
	{
		Pane.B();
		ErrorView = null;
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
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
			Uri resourceLocator = new Uri(XC.A(1156), UriKind.Relative);
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
	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
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
					scroller = (ScrollViewer)target;
					return;
				}
			}
		}
		this.m_A = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 2)
		{
			((Button)target).Click += LoadWorkspace;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
