using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Xml;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Model;

[DesignerGenerated]
public sealed class wpfSummaryStats : System.Windows.Window, IComponentConnector, IStyleConnector
{
	public ObservableCollection<Statistic> SummaryStats;

	[CompilerGenerated]
	[AccessedThroughProperty("lvStats")]
	private ListView m_A;

	[AccessedThroughProperty("btnUp")]
	[CompilerGenerated]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnDown")]
	private Button B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnRename")]
	private Button C;

	[CompilerGenerated]
	[AccessedThroughProperty("chkIgnoreCurrency")]
	private System.Windows.Controls.CheckBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private Button D;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private Button E;

	private bool m_A;

	internal virtual ListView lvStats
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
			SelectionChangedEventHandler value2 = lvStats_SelectionChanged;
			ListView listView = this.m_A;
			if (listView != null)
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
				listView.SelectionChanged -= value2;
			}
			this.m_A = value;
			listView = this.m_A;
			if (listView == null)
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
				listView.SelectionChanged += value2;
				return;
			}
		}
	}

	internal virtual Button btnUp
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
			RoutedEventHandler value2 = btnUp_Click;
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

	internal virtual Button btnDown
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
			RoutedEventHandler value2 = btnDown_Click;
			Button button = B;
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
				switch (4)
				{
				case 0:
					continue;
				}
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnRename
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
			RoutedEventHandler value2 = btnRename_Click;
			Button button = C;
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
				button.Click -= value2;
			}
			C = value;
			button = C;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkIgnoreCurrency
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

	internal virtual Button btnOk
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
			RoutedEventHandler value2 = btnOk_Click;
			Button button = D;
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

	internal virtual Button btnCancel
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
			E = value;
		}
	}

	public wpfSummaryStats()
	{
		base.Loaded += wpfSummaryStats_Loaded;
		base.Closing += wpfSummaryStats_Closing;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void wpfSummaryStats_Loaded(object sender, RoutedEventArgs e)
	{
		XmlDocument settingsXml = KH.A.SettingsXml;
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		bool flag = application.LanguageSettings.get_LanguageID(MsoAppLanguageID.msoLanguageIDUI) != 1033;
		SummaryStats = new ObservableCollection<Statistic>();
		Name name = default(Name);
		XmlNodeList xmlNodeList;
		try
		{
			if (flag)
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
				name = application.ActiveWorkbook.Names.Add(VH.A(94040), VH.A(94065), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			xmlNodeList = settingsXml.SelectNodes(VH.A(94072));
			_ = lvStats.Items;
			foreach (XmlNode item in xmlNodeList)
			{
				string value = item.Attributes[VH.A(94085)].Value;
				string strFxnLocal;
				if (!flag)
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
					strFxnLocal = value;
				}
				else
				{
					try
					{
						if (Operators.CompareString(value, VH.A(94102), TextCompare: false) != 0)
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
							if (Operators.CompareString(value, VH.A(94127), TextCompare: false) != 0)
							{
								name.RefersTo = VH.A(48936) + value + VH.A(94175);
								goto IL_01f5;
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
						name.RefersTo = VH.A(48936) + value + VH.A(94152);
						goto IL_01f5;
						IL_01f5:
						string text = name.RefersToLocal.ToString();
						strFxnLocal = Strings.Mid(text, 2, checked(text.IndexOf(VH.A(39848)) - 1));
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						strFxnLocal = value;
						ProjectData.ClearProjectError();
					}
				}
				SummaryStats.Add(new Statistic(Conversions.ToBoolean(item.Attributes[VH.A(94190)].Value), item.Attributes[VH.A(94205)].Value, value, strFxnLocal));
			}
			_ = null;
			chkIgnoreCurrency.IsChecked = Conversions.ToBoolean(settingsXml.SelectSingleNode(VH.A(94216)).InnerText);
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.ErrorMessage(ex4.Message);
			Close();
			ProjectData.ClearProjectError();
		}
		lvStats.ItemsSource = SummaryStats;
		if (name != null)
		{
			name.Delete();
			name = null;
		}
		settingsXml = null;
		xmlNodeList = null;
		application = null;
	}

	private void wpfSummaryStats_Closing(object sender, CancelEventArgs e)
	{
		if (base.DialogResult != true)
		{
			return;
		}
		int num = 0;
		XmlDocument settingsXml;
		try
		{
			settingsXml = KH.A.SettingsXml;
			settingsXml.SelectSingleNode(VH.A(94216)).InnerText = chkIgnoreCurrency.IsChecked.ToString();
			XmlNode xmlNode = settingsXml.SelectSingleNode(VH.A(94249));
			IEnumerator<Statistic> enumerator = default(IEnumerator<Statistic>);
			try
			{
				enumerator = SummaryStats.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Statistic current = enumerator.Current;
					XmlNode xmlNode2 = xmlNode.ChildNodes[num];
					xmlNode2.Attributes[VH.A(94205)].Value = current.Label;
					xmlNode2.Attributes[VH.A(94085)].Value = current.FunctionEnglish;
					xmlNode2.Attributes[VH.A(94190)].Value = current.IsChecked.ToString();
					_ = null;
					num = checked(num + 1);
				}
			}
			finally
			{
				if (enumerator != null)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						enumerator.Dispose();
						break;
					}
				}
			}
			xmlNode = null;
			KH.A.SaveSettings(settingsXml);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
		settingsXml = null;
	}

	private void lvStats_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (lvStats.SelectedIndex <= -1)
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
			btnRename.IsEnabled = true;
			int selectedIndex = lvStats.SelectedIndex;
			if (selectedIndex == 0)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						btnUp.IsEnabled = false;
						btnDown.IsEnabled = true;
						return;
					}
				}
			}
			if (selectedIndex == checked(SummaryStats.Count - 1))
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						btnUp.IsEnabled = true;
						btnDown.IsEnabled = false;
						return;
					}
				}
			}
			btnUp.IsEnabled = true;
			btnDown.IsEnabled = true;
			return;
		}
	}

	private void FunctionToggled(object sender, RoutedEventArgs e)
	{
		lvStats.SelectedItem = RuntimeHelpers.GetObjectValue(((System.Windows.Controls.CheckBox)sender).DataContext);
	}

	private void btnUp_Click(object sender, RoutedEventArgs e)
	{
		A(-1);
	}

	private void btnDown_Click(object sender, RoutedEventArgs e)
	{
		A(1);
	}

	private void A(int A)
	{
		ListView listView = lvStats;
		int selectedIndex = listView.SelectedIndex;
		checked
		{
			Statistic value;
			Statistic value2;
			try
			{
				ObservableCollection<Statistic> summaryStats = SummaryStats;
				value = summaryStats[selectedIndex];
				value2 = summaryStats[selectedIndex + A];
				summaryStats[selectedIndex] = value2;
				summaryStats[selectedIndex + A] = value;
				_ = null;
				listView.SelectedIndex = selectedIndex + A;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			listView = null;
			value = null;
			value2 = null;
		}
	}

	private void btnRename_Click(object sender, RoutedEventArgs e)
	{
		string label = SummaryStats[lvStats.SelectedIndex].Label;
		string text = Forms.InputBox(System.Windows.Window.GetWindow(this), VH.A(94264), VH.A(94297), label);
		if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
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
			if (text.Length <= 0)
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
				if (Operators.CompareString(text, label, TextCompare: false) == 0)
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
					SummaryStats[lvStats.SelectedIndex].Label = text;
					return;
				}
			}
		}
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = true;
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
			Uri resourceLocator = new Uri(VH.A(94318), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
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
					lvStats = (ListView)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnUp = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnDown = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnRename = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			chkIgnoreCurrency = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 7)
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
		if (connectionId == 8)
		{
			while (true)
			{
				switch (3)
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

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 2)
		{
			((System.Windows.Controls.CheckBox)target).Checked += FunctionToggled;
			((System.Windows.Controls.CheckBox)target).Unchecked += FunctionToggled;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
