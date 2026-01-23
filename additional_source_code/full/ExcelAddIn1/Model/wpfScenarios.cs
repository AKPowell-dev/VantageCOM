using System;
using System.CodeDom.Compiler;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using A;
using MacabacusMacros.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Model;

[DesignerGenerated]
public sealed class wpfScenarios : Window, IComponentConnector
{
	public bool GlobalScenariosFound;

	[CompilerGenerated]
	[AccessedThroughProperty("radGlobal")]
	private RadioButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lvGlobal")]
	private ListBox m_A;

	[AccessedThroughProperty("btnAddGlobal")]
	[CompilerGenerated]
	private Button m_A;

	[AccessedThroughProperty("btnDeleteGlobal")]
	[CompilerGenerated]
	private Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnEditGlobal")]
	private Button m_C;

	[AccessedThroughProperty("radLocal")]
	[CompilerGenerated]
	private RadioButton m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("lvLocal")]
	private ListBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnAddLocal")]
	private Button D;

	[AccessedThroughProperty("btnDeleteLocal")]
	[CompilerGenerated]
	private Button E;

	[CompilerGenerated]
	[AccessedThroughProperty("btnEditLocal")]
	private Button F;

	[AccessedThroughProperty("btnOk")]
	[CompilerGenerated]
	private Button G;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private Button H;

	private bool m_A;

	internal virtual RadioButton radGlobal
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
			RoutedEventHandler value2 = ScopeChanged;
			RoutedEventHandler value3 = ScopeChanged;
			RadioButton radioButton = this.m_A;
			if (radioButton != null)
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
				radioButton.Checked -= value2;
				radioButton.Unchecked -= value3;
			}
			this.m_A = value;
			radioButton = this.m_A;
			if (radioButton == null)
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
				radioButton.Checked += value2;
				radioButton.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual ListBox lvGlobal
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

	internal virtual Button btnAddGlobal
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
			RoutedEventHandler value2 = AddGlobalScenario;
			Button button = this.m_A;
			if (button != null)
			{
				while (true)
				{
					switch (2)
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

	internal virtual Button btnDeleteGlobal
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
			RoutedEventHandler value2 = DeleteGlobalScenario;
			Button button = this.m_B;
			if (button != null)
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
				button.Click -= value2;
			}
			this.m_B = value;
			button = this.m_B;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnEditGlobal
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = EditGlobalScenario;
			Button button = this.m_C;
			if (button != null)
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
				button.Click -= value2;
			}
			this.m_C = value;
			button = this.m_C;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual RadioButton radLocal
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
			RoutedEventHandler value2 = ScopeChanged;
			RoutedEventHandler value3 = ScopeChanged;
			RadioButton radioButton = this.m_B;
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
				radioButton.Checked -= value2;
				radioButton.Unchecked -= value3;
			}
			this.m_B = value;
			radioButton = this.m_B;
			if (radioButton == null)
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
				radioButton.Checked += value2;
				radioButton.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual ListBox lvLocal
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

	internal virtual Button btnAddLocal
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = AddLocalScenario;
			Button button = D;
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
			D = value;
			button = D;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnDeleteLocal
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = DeleteLocalScenario;
			Button button = E;
			if (button != null)
			{
				while (true)
				{
					switch (2)
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
			E = value;
			button = E;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnEditLocal
	{
		[CompilerGenerated]
		get
		{
			return F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = EditLocalScenario;
			Button button = F;
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
			F = value;
			button = F;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnOk
	{
		[CompilerGenerated]
		get
		{
			return G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnOk_Click;
			Button button = G;
			if (button != null)
			{
				while (true)
				{
					switch (2)
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
			G = value;
			button = G;
			if (button == null)
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
			return H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			H = value;
		}
	}

	public wpfScenarios()
	{
		base.Closing += wpfScenarios_Closing;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void wpfScenarios_Closing(object sender, CancelEventArgs e)
	{
		if (base.DialogResult != true)
		{
			return;
		}
		bool? isChecked = radGlobal.IsChecked;
		if (isChecked.HasValue)
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
			if (isChecked != true)
			{
				goto IL_0085;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		if (lvGlobal.Items.Count == 0)
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
			if (isChecked.HasValue)
			{
				goto IL_00dc;
			}
		}
		goto IL_0085;
		IL_00dc:
		Forms.WarningMessage(VH.A(93800));
		e.Cancel = true;
		return;
		IL_0085:
		isChecked = radLocal.IsChecked;
		if (isChecked.HasValue)
		{
			if (isChecked != true)
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
				break;
			}
		}
		if (lvLocal.Items.Count != 0 || !isChecked.HasValue)
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
			break;
		}
		goto IL_00dc;
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = true;
	}

	private void ScopeChanged(object sender, RoutedEventArgs e)
	{
		if (((RadioButton)sender).IsChecked != true)
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
			if (radGlobal.IsChecked == true)
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
				A(A: false);
				if (GlobalScenariosFound)
				{
					B(A: false);
				}
				else
				{
					B(A: true);
				}
			}
			else
			{
				A(A: true);
				B(A: false);
			}
			lvGlobal.SelectedIndex = -1;
			lvLocal.SelectedIndex = -1;
			return;
		}
	}

	private void A(bool A)
	{
		lvGlobal.IsEnabled = !A;
		lvLocal.IsEnabled = A;
		btnAddLocal.IsEnabled = A;
		btnDeleteLocal.IsEnabled = A;
		btnEditLocal.IsEnabled = A;
	}

	private void B(bool A)
	{
		btnAddGlobal.IsEnabled = A;
		btnDeleteGlobal.IsEnabled = A;
		btnEditGlobal.IsEnabled = A;
	}

	private void AddGlobalScenario(object sender, RoutedEventArgs e)
	{
		A(lvGlobal);
	}

	private void DeleteGlobalScenario(object sender, RoutedEventArgs e)
	{
		B(lvGlobal);
	}

	private void EditGlobalScenario(object sender, RoutedEventArgs e)
	{
		C(lvGlobal);
	}

	private void AddLocalScenario(object sender, RoutedEventArgs e)
	{
		A(lvLocal);
	}

	private void DeleteLocalScenario(object sender, RoutedEventArgs e)
	{
		B(lvLocal);
	}

	private void EditLocalScenario(object sender, RoutedEventArgs e)
	{
		C(lvLocal);
	}

	private void A(ListBox A)
	{
		string text = this.A();
		if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
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
			this.A(A).Add(text);
			return;
		}
	}

	private void B(ListBox A)
	{
		if (A.SelectedIndex <= -1)
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
			this.A(A).RemoveAt(A.SelectedIndex);
			return;
		}
	}

	private void C(ListBox A)
	{
		if (A.SelectedIndex <= -1)
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
			string text = this.A();
			if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
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
				this.A(A)[A.SelectedIndex] = text;
				return;
			}
		}
	}

	private ObservableCollection<string> A(ListBox A)
	{
		return (ObservableCollection<string>)A.ItemsSource;
	}

	private string A()
	{
		string text = Forms.InputBox(Window.GetWindow(this), VH.A(93847), VH.A(93866), "");
		if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
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
			if (text.Length > 0)
			{
				return text;
			}
		}
		return string.Empty;
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
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_A = true;
			Uri resourceLocator = new Uri(VH.A(93919), UriKind.Relative);
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
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					radGlobal = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			lvGlobal = (ListBox)target;
			return;
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
					btnAddGlobal = (Button)target;
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
					btnDeleteGlobal = (Button)target;
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
					btnEditGlobal = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			radLocal = (RadioButton)target;
			return;
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					lvLocal = (ListBox)target;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnAddLocal = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnDeleteLocal = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnEditLocal = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnOk = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnCancel = (Button)target;
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
}
