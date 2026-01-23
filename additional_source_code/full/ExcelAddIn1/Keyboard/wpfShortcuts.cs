using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Config;
using MacabacusMacros.Config.Settings;
using MacabacusMacros.ShortcutManager;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Keyboard;

[DesignerGenerated]
public sealed class wpfShortcuts : System.Windows.Controls.UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[CompilerGenerated]
	internal sealed class GG
	{
		public ShortcutItem A;

		public wpfShortcuts A;

		[SpecialName]
		internal void A()
		{
			this.A.lvShortcuts.SelectionChanged -= this.A.ShortcutSelectionChanged;
			this.A.IsSelected = true;
			this.A.D();
			this.A.lvShortcuts.SelectionChanged += this.A.ShortcutSelectionChanged;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private static FieldInfo m_A;

	private bool m_A;

	private bool m_B;

	private GridViewColumnHeader m_A;

	private SortAdorner m_A;

	private ScrollViewer m_A;

	private ICollectionView m_A;

	[CompilerGenerated]
	private ObservableCollection<ShortcutItem> m_A;

	private ObservableCollection<string> m_A;

	private Thickness m_A;

	private Visibility m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("ThisWindow")]
	private wpfShortcuts m_A;

	[AccessedThroughProperty("lvShortcuts")]
	[CompilerGenerated]
	private System.Windows.Controls.ListView m_A;

	[AccessedThroughProperty("gvShortcuts")]
	[CompilerGenerated]
	private GridView m_A;

	[AccessedThroughProperty("btnReset")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnDefaults")]
	private System.Windows.Controls.Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClear")]
	private System.Windows.Controls.Button m_C;

	private bool m_C;

	public ICollectionView SourceCollection
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(36261));
		}
	}

	private ObservableCollection<ShortcutItem> AllItems
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public ObservableCollection<string> AllowedKeys
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(161406));
		}
	}

	public Thickness AdornerPadding
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(161429));
		}
	}

	public Visibility EditVisibility
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(161458));
		}
	}

	internal virtual wpfShortcuts ThisWindow
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

	internal virtual System.Windows.Controls.ListView lvShortcuts
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
			SizeChangedEventHandler value2 = ListViewSizeChanged;
			System.Windows.Controls.ListView listView = this.m_A;
			if (listView != null)
			{
				listView.SizeChanged -= value2;
			}
			this.m_A = value;
			listView = this.m_A;
			if (listView == null)
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
				listView.SizeChanged += value2;
				return;
			}
		}
	}

	internal virtual GridView gvShortcuts
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

	internal virtual System.Windows.Controls.Button btnReset
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
			RoutedEventHandler value2 = ResetShortcut;
			System.Windows.Controls.Button button = this.m_A;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.m_A = value;
			button = this.m_A;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnDefaults
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
			RoutedEventHandler value2 = ResetAll;
			System.Windows.Controls.Button button = this.m_B;
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

	internal virtual System.Windows.Controls.Button btnClear
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
			RoutedEventHandler value2 = ClearAll;
			System.Windows.Controls.Button button = this.m_C;
			if (button != null)
			{
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
				switch (3)
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

	public wpfShortcuts()
	{
		base.Loaded += wpfShortcuts_Loaded;
		this.m_A = true;
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
		this.m_A = Visibility.Visible;
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
			switch (4)
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

	private static void A()
	{
		wpfShortcuts.m_A = typeof(SystemParameters).GetField(VH.A(42651), BindingFlags.Static | BindingFlags.NonPublic);
		B();
		SystemParameters.StaticPropertyChanged += SystemParameters_StaticPropertyChanged;
	}

	private static void SystemParameters_StaticPropertyChanged(object sender, PropertyChangedEventArgs e)
	{
		B();
	}

	private static void B()
	{
		if (!SystemParameters.MenuDropAlignment)
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
			if ((object)wpfShortcuts.m_A != null)
			{
				wpfShortcuts.m_A.SetValue(null, false);
			}
			return;
		}
	}

	private void wpfShortcuts_Loaded(object sender, RoutedEventArgs e)
	{
		clsPanes.EliminateTaskPaneFlicker(this);
		this.m_B = SharedSettings.IsSettingEditable(Constants.XML_SHORTCUTS);
		if (this.m_B)
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
			EditVisibility = Visibility.Visible;
			System.Windows.Controls.ListView listView = lvShortcuts;
			listView.PreviewKeyDown += SpacebarToggle;
			listView.KeyDown += ModifyKeystroke;
			listView.KeyUp += CtrlShift0;
			_ = null;
		}
		else
		{
			EditVisibility = Visibility.Collapsed;
			B(VH.A(161487));
		}
		this.m_A = (ScrollViewer)Forms.GetScrollViewer((DependencyObject)lvShortcuts);
		this.m_A.ScrollChanged += ScrollChanged;
		PopulateList();
	}

	public void PopulateList(XmlDocument xmlDoc = null)
	{
		//IL_015f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0183: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_0232: Unknown result type (might be due to invalid IL or missing references)
		//IL_023b: Unknown result type (might be due to invalid IL or missing references)
		//IL_026d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0274: Unknown result type (might be due to invalid IL or missing references)
		//IL_027b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0285: Expected O, but got Unknown
		List<string> list = new List<string>();
		XmlNodeList xmlNodeList = KH.A.SettingsXml.DocumentElement.SelectNodes(VH.A(161606));
		XmlDocument xmlDocument = new XmlDocument();
		new List<System.Windows.Controls.ListViewItem>();
		if (xmlDoc == null)
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
			xmlDoc = KH.A.SettingsXml;
		}
		xmlDocument.LoadXml(J.DefaultShortcuts);
		AllowedKeys = Base.ComboBoxKeys();
		AllItems = new ObservableCollection<ShortcutItem>();
		AdornerPadding = new Thickness(0.0);
		this.m_A = true;
		if (this.m_A != null)
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
			if (this.m_A != null)
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
				AdornerLayer.GetAdornerLayer(this.m_A).Remove((Adorner)(object)this.m_A);
				this.m_A = null;
				this.m_A = null;
			}
		}
		XmlNodeList xmlNodeList2 = xmlDocument.DocumentElement.SelectNodes(VH.A(161643));
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = xmlNodeList2.GetEnumerator();
			while (enumerator.MoveNext())
			{
				XmlNode xmlNode = (XmlNode)enumerator.Current;
				try
				{
					string value = xmlNode.Attributes[VH.A(161660)].Value;
					if (!Shortcuts.ShortcutsDictionary.TryGetValue(value, out var value2))
					{
						continue;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						string category = value2.Category;
						if (!list.Contains(category))
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
							list.Add(category);
						}
						if (value2.FriendlyName.Length <= 0)
						{
							break;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							if (value.Length <= 0)
							{
								break;
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								string text = value2.FriendlyName;
								if (text.StartsWith(VH.A(161671)))
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
									text = xmlNodeList.Item(checked(Conversions.ToInteger(Regex.Replace(text, VH.A(161696), "")) - 1)).Attributes[VH.A(67336)].Value;
								}
								AllItems.Add(new ShortcutItem(value2.Category, value, text, value2.Utility, xmlDoc.SelectSingleNode(Shortcuts.ShortcutXpath(value)).Attributes[VH.A(161707)].Value, value2.DefaultKey, value2.Description));
								break;
							}
							break;
						}
						break;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_02a4;
				}
				continue;
				end_IL_02a4:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		xmlNodeList2 = null;
		lvShortcuts.SelectionChanged -= ShortcutSelectionChanged;
		SourceCollection = CollectionViewSource.GetDefaultView(AllItems);
		SourceCollection.GroupDescriptions.Add(new PropertyGroupDescription(VH.A(161726)));
		lvShortcuts.SelectionChanged += ShortcutSelectionChanged;
		this.m_A = false;
		base.Dispatcher.BeginInvoke((global::A.A)([SpecialName] () =>
		{
			UpdateColumnWidths(lvShortcuts);
		}), DispatcherPriority.Normal);
		xmlNodeList = null;
		list = null;
		xmlDoc = null;
	}

	private void LastKeyChanged(object sender, SelectionChangedEventArgs e)
	{
		//IL_0048: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Expected O, but got Unknown
		if (!base.IsLoaded)
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
				System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
				if (comboBox.IsLoaded)
				{
					ShortcutItem val = (ShortcutItem)comboBox.DataContext;
					if (val.LastKey.Length > 0)
					{
						A(val);
					}
					val = null;
				}
				comboBox = null;
				return;
			}
		}
	}

	private void ModifyKeystroke(object sender, System.Windows.Input.KeyEventArgs e)
	{
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004a: Expected O, but got Unknown
		if (lvShortcuts.SelectedItems.Count <= 0)
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
			ShortcutItem val = (ShortcutItem)lvShortcuts.SelectedItems[0];
			if (val.IsEnabled)
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
				bool flag = System.Windows.Input.Keyboard.Modifiers.HasFlag(ModifierKeys.Control);
				bool flag2 = System.Windows.Input.Keyboard.Modifiers.HasFlag(ModifierKeys.Alt);
				bool flag3 = System.Windows.Input.Keyboard.Modifiers.HasFlag(ModifierKeys.Shift);
				this.m_A = true;
				ShortcutItem obj = val;
				obj.Clear();
				obj.ModifierCtrl = flag;
				obj.ModifierAlt = flag2;
				int modifierShift;
				if (flag3)
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
					modifierShift = ((flag || flag2) ? 1 : 0);
				}
				else
				{
					modifierShift = 0;
				}
				obj.ModifierShift = (byte)modifierShift != 0;
				_ = null;
				if (flag)
				{
					goto IL_010f;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				if (flag2)
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
					if (flag3)
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
						goto IL_010f;
					}
				}
				goto IL_0174;
			}
			goto IL_017b;
			IL_017b:
			val = null;
			return;
			IL_0174:
			this.m_A = false;
			goto IL_017b;
			IL_010f:
			Key key = e.Key;
			if ((uint)(key - 116) > 5u && key != Key.System)
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
				string keyFromPress = Base.GetKeyFromPress(e);
				if (!AllowedKeys.Contains(keyFromPress))
				{
					AllowedKeys.Add(keyFromPress);
				}
				val.LastKey = keyFromPress;
				A(val);
				e.Handled = true;
			}
			goto IL_0174;
		}
	}

	private void CtrlShift0(object sender, System.Windows.Input.KeyEventArgs e)
	{
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		//IL_0020: Expected O, but got Unknown
		if (Base.IsCtrlShift0(e, (ShortcutItem)lvShortcuts.SelectedItems[0]))
		{
			D();
		}
	}

	private void A(ShortcutItem A)
	{
		if (this.A(A, B: false))
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
					B(A);
					C();
					return;
				}
			}
		}
		this.m_A = true;
		A.Revert();
		this.m_A = false;
	}

	private bool A(ShortcutItem A, bool B)
	{
		if (A.LastKey.Length != 0)
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
			if (!A.ModifierCtrl && !A.ModifierAlt)
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
				if (!A.ModifierShift)
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
					goto IL_0056;
				}
			}
			string text = Base.BuildKeystroke(A);
			if (Operators.CompareString(text, VH.A(161780), TextCompare: false) == 0)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						Base.CtrlShift0Message();
						return false;
					}
				}
			}
			if (this.A(text))
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
				if (!Base.ConfirmOverrideNative())
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							return false;
						}
					}
				}
			}
			using (IEnumerator<ShortcutItem> enumerator = AllItems.GetEnumerator())
			{
				while (true)
				{
					if (enumerator.MoveNext())
					{
						ShortcutItem current = enumerator.Current;
						if (Operators.CompareString(current.Keystroke, text, TextCompare: false) != 0)
						{
							continue;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							break;
						}
						if (current == A)
						{
							continue;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							if (Base.ConfirmOverwrite(current))
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									current.IsEnabled = false;
									C(current);
									break;
								}
								break;
							}
							return false;
						}
						break;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_0152;
						}
						continue;
						end_IL_0152:
						break;
					}
					break;
				}
			}
			return true;
		}
		goto IL_0056;
		IL_0056:
		if (!B)
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
			this.B(VH.A(161743));
		}
		return false;
	}

	private bool A(string A)
	{
		string text = A.ToLower();
		uint num = TH.A(text);
		if (num <= 698900162)
		{
			if (num <= 421688663)
			{
				if (num <= 363347782)
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
					if (num <= 203579616)
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
						if (num != 104072603)
						{
							if (num != 203579616)
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
							}
							else
							{
								if (Operators.CompareString(text, VH.A(162424), TextCompare: false) == 0)
								{
									goto IL_0ac4;
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
							}
						}
						else if (Operators.CompareString(text, VH.A(162310), TextCompare: false) == 0)
						{
							goto IL_0ac0;
						}
					}
					else if (num != 329792544)
					{
						if (num != 337800568)
						{
							if (num != 363347782)
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
							}
							else
							{
								if (Operators.CompareString(text, VH.A(161818), TextCompare: false) == 0)
								{
									goto IL_0aac;
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
						}
						else
						{
							if (Operators.CompareString(text, VH.A(162382), TextCompare: false) == 0)
							{
								goto IL_0ac4;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								break;
							}
						}
					}
					else
					{
						if (Operators.CompareString(text, VH.A(161987), TextCompare: false) == 0)
						{
							goto IL_0ab4;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							break;
						}
					}
				}
				else if (num <= 396903020)
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
					if (num != 380125401)
					{
						if (num != 388133425)
						{
							if (num != 396903020)
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
							}
							else if (Operators.CompareString(text, VH.A(161935), TextCompare: false) == 0)
							{
								goto IL_0ab4;
							}
						}
						else
						{
							if (Operators.CompareString(text, VH.A(162387), TextCompare: false) == 0)
							{
								goto IL_0ac4;
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
					}
					else
					{
						if (Operators.CompareString(text, VH.A(161831), TextCompare: false) == 0)
						{
							goto IL_0aac;
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
					}
				}
				else if (num != 404911044)
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
					if (num != 413680639)
					{
						if (num != 421688663)
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
						}
						else
						{
							if (Operators.CompareString(text, VH.A(162392), TextCompare: false) == 0)
							{
								goto IL_0ac4;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								break;
							}
						}
					}
					else
					{
						if (Operators.CompareString(text, VH.A(162013), TextCompare: false) == 0)
						{
							goto IL_0ab8;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							break;
						}
					}
				}
				else
				{
					if (Operators.CompareString(text, VH.A(162414), TextCompare: false) == 0)
					{
						goto IL_0ac4;
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
				}
			}
			else if (num <= 547901591)
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
				if (num <= 447235877)
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
					if (num != 438466282)
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
						if (num != 447235877)
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
						}
						else if (Operators.CompareString(text, VH.A(162039), TextCompare: false) == 0)
						{
							goto IL_0ab8;
						}
					}
					else if (Operators.CompareString(text, VH.A(162419), TextCompare: false) == 0)
					{
						goto IL_0ac4;
					}
				}
				else if (num != 497568734)
				{
					if (num != 514346353)
					{
						if (num == 547901591 && Operators.CompareString(text, VH.A(161857), TextCompare: false) == 0)
						{
							goto IL_0ab0;
						}
					}
					else if (Operators.CompareString(text, VH.A(161948), TextCompare: false) == 0)
					{
						goto IL_0ab4;
					}
				}
				else
				{
					if (Operators.CompareString(text, VH.A(161922), TextCompare: false) == 0)
					{
						goto IL_0ab4;
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
			}
			else if (num <= 581456829)
			{
				if (num != 564679210)
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
					if (num != 581318465)
					{
						if (num != 581456829)
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
						}
						else
						{
							if (Operators.CompareString(text, VH.A(161870), TextCompare: false) == 0)
							{
								goto IL_0ab0;
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
					}
					else
					{
						if (Operators.CompareString(text, VH.A(162293), TextCompare: false) == 0)
						{
							goto IL_0ac0;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							break;
						}
					}
				}
				else
				{
					if (Operators.CompareString(text, VH.A(162000), TextCompare: false) == 0)
					{
						goto IL_0ab8;
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
			}
			else if (num != 665344924)
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
				if (num != 682122543)
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
					if (num != 698900162)
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
					}
					else
					{
						if (Operators.CompareString(text, VH.A(162052), TextCompare: false) == 0)
						{
							goto IL_0ab8;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							break;
						}
					}
				}
				else if (Operators.CompareString(text, VH.A(162065), TextCompare: false) == 0)
				{
					goto IL_0ab8;
				}
			}
			else if (Operators.CompareString(text, VH.A(161896), TextCompare: false) == 0)
			{
				goto IL_0ab0;
			}
		}
		else if (num <= 1477700177)
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
			if (num <= 799565876)
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
				if (num <= 749233019)
				{
					if (num != 732455400)
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
						if (num != 749233019)
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
						}
						else
						{
							if (Operators.CompareString(text, VH.A(162026), TextCompare: false) == 0)
							{
								goto IL_0ab8;
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
						}
					}
					else if (Operators.CompareString(text, VH.A(161974), TextCompare: false) == 0)
					{
						goto IL_0ab4;
					}
				}
				else if (num != 766010638)
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
					if (num != 782788257)
					{
						if (num != 799565876)
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
						}
						else
						{
							if (Operators.CompareString(text, VH.A(161909), TextCompare: false) == 0)
							{
								goto IL_0ab4;
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
					}
					else
					{
						if (Operators.CompareString(text, VH.A(161844), TextCompare: false) == 0)
						{
							goto IL_0aac;
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
				}
				else if (Operators.CompareString(text, VH.A(161883), TextCompare: false) == 0)
				{
					goto IL_0ab0;
				}
			}
			else if (num <= 872303894)
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
				if (num != 826951197)
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
					if (num != 849898733)
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
						if (num != 872303894)
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
						}
						else
						{
							if (Operators.CompareString(text, VH.A(162131), TextCompare: false) == 0)
							{
								goto IL_0abc;
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
					}
					else
					{
						if (Operators.CompareString(text, VH.A(161805), TextCompare: false) == 0)
						{
							goto IL_0aac;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							break;
						}
					}
				}
				else
				{
					if (Operators.CompareString(text, VH.A(162112), TextCompare: false) == 0)
					{
						goto IL_0abc;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						break;
					}
				}
			}
			else if (num != 1464394840)
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
				if (num != 1466661412)
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
					if (num != 1477700177)
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
					}
					else
					{
						if (Operators.CompareString(text, VH.A(162097), TextCompare: false) == 0)
						{
							goto IL_0abc;
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
					}
				}
				else if (Operators.CompareString(text, VH.A(162152), TextCompare: false) == 0)
				{
					goto IL_0abc;
				}
			}
			else
			{
				if (Operators.CompareString(text, VH.A(162078), TextCompare: false) == 0)
				{
					goto IL_0abc;
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					break;
				}
			}
		}
		else if (num <= 3160966811u)
		{
			if (num <= 2355079881u)
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
				if (num != 1566262289)
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
					if (num != 2355079881u)
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
					}
					else
					{
						if (Operators.CompareString(text, VH.A(162210), TextCompare: false) == 0)
						{
							goto IL_0abc;
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							break;
						}
					}
				}
				else
				{
					if (Operators.CompareString(text, VH.A(162397), TextCompare: false) == 0)
					{
						goto IL_0ac4;
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
				}
			}
			else if (num != 2635452357u)
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
				if (num != 2908604339u)
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
					if (num != 3160966811u)
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
					}
					else if (Operators.CompareString(text, VH.A(162274), TextCompare: false) == 0)
					{
						goto IL_0ac0;
					}
				}
				else
				{
					if (Operators.CompareString(text, VH.A(162348), TextCompare: false) == 0)
					{
						goto IL_0ac0;
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
			}
			else
			{
				if (Operators.CompareString(text, VH.A(162183), TextCompare: false) == 0)
				{
					goto IL_0abc;
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					break;
				}
			}
		}
		else if (num <= 3998789804u)
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
			if (num != 3308457076u)
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
				if (num != 3351539004u)
				{
					if (num != 3998789804u)
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
					}
					else if (Operators.CompareString(text, VH.A(162329), TextCompare: false) == 0)
					{
						goto IL_0ac0;
					}
				}
				else
				{
					if (Operators.CompareString(text, VH.A(162429), TextCompare: false) == 0)
					{
						goto IL_0ac4;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						break;
					}
				}
			}
			else
			{
				if (Operators.CompareString(text, VH.A(162365), TextCompare: false) == 0)
				{
					goto IL_0ac0;
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					break;
				}
			}
		}
		else if (num != 4037543175u)
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
			if (num != 4081400426u)
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
				if (num == 4231138174u)
				{
					if (Operators.CompareString(text, VH.A(162444), TextCompare: false) == 0)
					{
						goto IL_0ac4;
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
			}
			else if (Operators.CompareString(text, VH.A(162241), TextCompare: false) == 0)
			{
				goto IL_0abc;
			}
		}
		else
		{
			if (Operators.CompareString(text, VH.A(161961), TextCompare: false) == 0)
			{
				goto IL_0ab4;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		return false;
		IL_0aac:
		return true;
		IL_0ac4:
		return true;
		IL_0ab8:
		return true;
		IL_0ab4:
		return true;
		IL_0ab0:
		return true;
		IL_0ac0:
		return true;
		IL_0abc:
		return true;
	}

	private void B(ShortcutItem A)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		string text = Base.BuildKeystroke(A);
		try
		{
			XmlDocument settingsXml = KH.A.SettingsXml;
			XmlNode xmlNode = this.A(settingsXml, A);
			if (Operators.CompareString(xmlNode.Attributes[VH.A(161707)].Value, "", TextCompare: false) != 0)
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
				application.OnKey(Shortcuts.ConvertKeystroke(xmlNode.Attributes[VH.A(161707)].Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			application.OnKey(Shortcuts.ConvertKeystroke(text), xmlNode.Attributes[VH.A(161660)].Value);
			xmlNode.Attributes[VH.A(161707)].Value = text;
			xmlNode = null;
			this.A(settingsXml);
			A.Keystroke = text;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			C(VH.A(162451) + text + VH.A(162589) + ex2.Message);
			ProjectData.ClearProjectError();
		}
		finally
		{
			XmlDocument settingsXml = null;
		}
		D();
		application = null;
	}

	private void A(XmlDocument A)
	{
		if (!this.m_B)
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
			KH.A.SaveSettings(A);
			Shortcuts.dictLookup2 = null;
			KH.A.Invalidate();
			return;
		}
	}

	private void SelectOnMouseDown(object sender, MouseButtonEventArgs e)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Expected O, but got Unknown
		ShortcutItem val = (ShortcutItem)((System.Windows.Controls.Control)sender).DataContext;
		if (!val.IsSelected)
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
			val.IsSelected = true;
			D();
			e.Handled = true;
		}
		val = null;
	}

	private void ShortcutSelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		//IL_0063: Unknown result type (might be due to invalid IL or missing references)
		//IL_006d: Expected O, but got Unknown
		if (e.OriginalSource != lvShortcuts)
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
			if (e.RemovedItems.Count > 0)
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
				ShortcutItem A = (ShortcutItem)e.RemovedItems[0];
				if (A.IsDirty)
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
					if (A.IsEnabled)
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
						if (!this.A(A, B: true))
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
							if (System.Windows.Forms.MessageBox.Show(VH.A(162602) + A.FriendlyName + VH.A(162699), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
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
								this.m_A = true;
								A.Revert();
								this.m_A = false;
								A = null;
							}
							else
							{
								base.Dispatcher.BeginInvoke((global::A.A)([SpecialName] () =>
								{
									lvShortcuts.SelectionChanged -= ShortcutSelectionChanged;
									A.IsSelected = true;
									D();
									lvShortcuts.SelectionChanged += ShortcutSelectionChanged;
								}), DispatcherPriority.Normal);
							}
						}
						else
						{
							A = null;
						}
					}
				}
			}
			if (lvShortcuts.SelectedItems.Count <= 0)
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
				btnReset.IsEnabled = this.m_B;
				return;
			}
		}
	}

	private void ShortcutCheckedChanged(object sender, RoutedEventArgs e)
	{
		//IL_0030: Unknown result type (might be due to invalid IL or missing references)
		//IL_0036: Expected O, but got Unknown
		if (!this.m_B)
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
			if (!this.m_A)
			{
				ShortcutItem val = (ShortcutItem)((System.Windows.Controls.CheckBox)sender).DataContext;
				this.m_A = true;
				val.Clear();
				if (!val.IsEnabled)
				{
					C(val);
				}
				this.m_A = false;
				val = null;
			}
			return;
		}
	}

	private void C(ShortcutItem A)
	{
		XmlDocument settingsXml = KH.A.SettingsXml;
		XmlAttribute xmlAttribute = this.A(settingsXml, A).Attributes[VH.A(161707)];
		if (xmlAttribute.Value.Length > 0)
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
			MH.A.Application.OnKey(Shortcuts.ConvertKeystroke(xmlAttribute.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		xmlAttribute.Value = "";
		xmlAttribute = null;
		this.A(settingsXml);
		settingsXml = null;
	}

	private void SpacebarToggle(object sender, System.Windows.Input.KeyEventArgs e)
	{
		Base.SpacebarToggle(lvShortcuts, e);
	}

	private void ResetShortcut(object sender, RoutedEventArgs e)
	{
		//IL_0042: Unknown result type (might be due to invalid IL or missing references)
		//IL_0048: Expected O, but got Unknown
		bool flag = false;
		if (lvShortcuts.SelectedItems.Count <= 0)
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
			ShortcutItem val = (ShortcutItem)lvShortcuts.SelectedItems[0];
			foreach (ShortcutItem allItem in AllItems)
			{
				if (Operators.CompareString(allItem.Keystroke, val.DefaultKeystroke, TextCompare: false) != 0)
				{
					continue;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					break;
				}
				if (allItem == val)
				{
					continue;
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					B(VH.A(162722) + allItem.FriendlyName + VH.A(64021));
					flag = true;
					break;
				}
				break;
			}
			if (!flag)
			{
				this.m_A = true;
				val.IsEnabled = true;
				this.m_A = false;
				val.Reset();
				B(val);
			}
			val = null;
			return;
		}
	}

	private void ResetAll(object sender, RoutedEventArgs e)
	{
		if (!Base.ConfirmResetAll())
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
			XmlDocument settingsXml = KH.A.SettingsXml;
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(J.DefaultShortcuts);
			XmlNode xmlNode = settingsXml.DocumentElement.SelectSingleNode(VH.A(162833));
			xmlNode.ParentNode.RemoveChild(xmlNode);
			settingsXml.DocumentElement.SelectSingleNode(VH.A(162882)).AppendChild(settingsXml.ImportNode(xmlDocument.DocumentElement, deep: true));
			xmlDocument = null;
			xmlNode = null;
			clsShortcuts.SanitizeShortcutsXml(ref settingsXml);
			Shortcuts.Remove();
			A(settingsXml);
			Shortcuts.Load();
			PopulateList(settingsXml);
			settingsXml = null;
			C();
			return;
		}
	}

	private void ClearAll(object sender, RoutedEventArgs e)
	{
		if (System.Windows.Forms.MessageBox.Show(VH.A(162901), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
		{
			return;
		}
		IEnumerator<ShortcutItem> enumerator = default(IEnumerator<ShortcutItem>);
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			this.m_A = true;
			XmlDocument settingsXml;
			try
			{
				settingsXml = KH.A.SettingsXml;
				try
				{
					enumerator = AllItems.GetEnumerator();
					while (enumerator.MoveNext())
					{
						ShortcutItem current = enumerator.Current;
						ShortcutItem val = current;
						if (val.IsEnabled)
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
							val.IsEnabled = false;
							XmlNode xmlNode = A(settingsXml, current);
							application.OnKey(Shortcuts.ConvertKeystroke(xmlNode.Attributes[VH.A(161707)].Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							xmlNode.Attributes[VH.A(161707)].Value = "";
							xmlNode = null;
						}
						val = null;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_0112;
						}
						continue;
						end_IL_0112:
						break;
					}
				}
				finally
				{
					if (enumerator != null)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							enumerator.Dispose();
							break;
						}
					}
				}
				A(settingsXml);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				C(ex2.Message);
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			this.m_A = false;
			application = null;
			settingsXml = null;
			return;
		}
	}

	private void SortColumn(object sender, RoutedEventArgs e)
	{
		//IL_00a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b0: Expected O, but got Unknown
		GridViewColumnHeader gridViewColumnHeader = (GridViewColumnHeader)sender;
		string text = gridViewColumnHeader.Tag.ToString();
		if (this.m_A != null)
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
			AdornerLayer.GetAdornerLayer(this.m_A).Remove((Adorner)(object)this.m_A);
			SourceCollection.SortDescriptions.Clear();
		}
		SourceCollection.GroupDescriptions.Clear();
		ListSortDirection listSortDirection = ListSortDirection.Descending;
		if (this.m_A == gridViewColumnHeader && this.m_A.Direction == listSortDirection)
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
			listSortDirection = ListSortDirection.Ascending;
		}
		this.m_A = gridViewColumnHeader;
		this.m_A = new SortAdorner((UIElement)this.m_A, listSortDirection);
		AdornerLayer.GetAdornerLayer(this.m_A).Add((Adorner)(object)this.m_A);
		SourceCollection.SortDescriptions.Add(new SortDescription(text, listSortDirection));
		if (Operators.CompareString(text, VH.A(163225), TextCompare: false) == 0)
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
			AdornerPadding = new Thickness(0.0, 0.0, 10.0, 0.0);
			Base.UpdateColumnWidth(gridViewColumnHeader.Column);
		}
		else
		{
			AdornerPadding = new Thickness(0.0);
			Base.UpdateColumnWidth(((GridView)lvShortcuts.View).Columns[2]);
		}
		UpdateColumnWidths(lvShortcuts);
		gridViewColumnHeader = null;
	}

	private void ListViewSizeChanged(object sender, SizeChangedEventArgs e)
	{
		if (!e.WidthChanged)
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
			UpdateColumnWidths(lvShortcuts, e);
			return;
		}
	}

	private void UpdateColumnWidths(System.Windows.Controls.ListView lv, SizeChangedEventArgs e = null)
	{
		if (this.m_A == null)
		{
			this.m_A = (ScrollViewer)Forms.GetScrollViewer((DependencyObject)lvShortcuts);
		}
		Base.UpdateFirstColumnWidth(this.m_A, gvShortcuts, e);
	}

	private void ScrollChanged(object sender, ScrollChangedEventArgs e)
	{
		if (this.m_A.VerticalOffset > 0.0)
		{
			base.Dispatcher.BeginInvoke((global::A.A)([SpecialName] () =>
			{
				Base.UpdateColumnWidth(gvShortcuts.Columns[1]);
			}), DispatcherPriority.ContextIdle);
		}
	}

	private void C()
	{
		Base.UpdateKeystrokeColumnWidth(lvShortcuts);
		UpdateColumnWidths(lvShortcuts);
	}

	private XmlNode A(XmlDocument A, ShortcutItem B)
	{
		return A.SelectSingleNode(Shortcuts.ShortcutXpath(B.MacroName));
	}

	private void D()
	{
		Base.FocusListViewItem(lvShortcuts);
	}

	private void A(GridView A)
	{
		Forms.AutoResizeGridView(A);
	}

	private void B(string A)
	{
		System.Windows.Forms.MessageBox.Show(A, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
	}

	private void C(string A)
	{
		System.Windows.Forms.MessageBox.Show(A, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (this.m_C)
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
			this.m_C = true;
			Uri resourceLocator = new Uri(VH.A(163240), UriKind.Relative);
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
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					ThisWindow = (wpfShortcuts)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					lvShortcuts = (System.Windows.Controls.ListView)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					gvShortcuts = (GridView)target;
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
					((GridViewColumnHeader)target).Click += SortColumn;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			((GridViewColumnHeader)target).Click += SortColumn;
			return;
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					((GridViewColumnHeader)target).Click += SortColumn;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			btnReset = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnDefaults = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnClear = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		this.m_C = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 5)
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
			((System.Windows.Controls.CheckBox)target).Checked += ShortcutCheckedChanged;
			((System.Windows.Controls.CheckBox)target).Unchecked += ShortcutCheckedChanged;
			((System.Windows.Controls.CheckBox)target).PreviewMouseDown += SelectOnMouseDown;
		}
		if (connectionId == 7)
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
			((System.Windows.Controls.CheckBox)target).PreviewMouseDown += SelectOnMouseDown;
		}
		if (connectionId == 8)
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
			((System.Windows.Controls.CheckBox)target).PreviewMouseDown += SelectOnMouseDown;
		}
		if (connectionId == 9)
		{
			((System.Windows.Controls.CheckBox)target).PreviewMouseDown += SelectOnMouseDown;
		}
		if (connectionId == 10)
		{
			((System.Windows.Controls.ComboBox)target).SelectionChanged += LastKeyChanged;
			((System.Windows.Controls.ComboBox)target).PreviewMouseDown += SelectOnMouseDown;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}

	[SpecialName]
	[CompilerGenerated]
	private void E()
	{
		UpdateColumnWidths(lvShortcuts);
	}

	[SpecialName]
	[CompilerGenerated]
	private void F()
	{
		Base.UpdateColumnWidth(gvShortcuts.Columns[1]);
	}
}
