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
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Keyboard;

[DesignerGenerated]
public sealed class wpfShortcuts : System.Windows.Controls.UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[CompilerGenerated]
	internal sealed class O
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

	[AccessedThroughProperty("ThisWindow")]
	[CompilerGenerated]
	private wpfShortcuts m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lvShortcuts")]
	private System.Windows.Controls.ListView m_A;

	[AccessedThroughProperty("gvShortcuts")]
	[CompilerGenerated]
	private GridView m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnReset")]
	private System.Windows.Controls.Button m_A;

	[AccessedThroughProperty("btnDefaults")]
	[CompilerGenerated]
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
			A(XC.A(3391));
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
			A(XC.A(3424));
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
			A(XC.A(3447));
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
			A(XC.A(3476));
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
			if (listView != null)
			{
				listView.SizeChanged += value2;
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
			this.m_B = value;
			button = this.m_B;
			if (button == null)
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
				button.Click -= value2;
			}
			this.m_C = value;
			button = this.m_C;
			if (button != null)
			{
				button.Click += value2;
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
			switch (3)
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
		wpfShortcuts.m_A = typeof(SystemParameters).GetField(XC.A(3354), BindingFlags.Static | BindingFlags.NonPublic);
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
			switch (1)
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
		this.m_B = SharedSettings.IsSettingEditable(Constants.XML_SHORTCUTS);
		if (this.m_B)
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
			Forms.WarningMessage(XC.A(3505));
		}
		this.m_A = (ScrollViewer)Forms.GetScrollViewer((DependencyObject)lvShortcuts);
		this.m_A.ScrollChanged += ScrollChanged;
		PopulateList();
	}

	public void PopulateList(XmlDocument xmlDoc = null)
	{
		//IL_0161: Unknown result type (might be due to invalid IL or missing references)
		//IL_0185: Unknown result type (might be due to invalid IL or missing references)
		//IL_01bb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0234: Unknown result type (might be due to invalid IL or missing references)
		//IL_023d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0275: Unknown result type (might be due to invalid IL or missing references)
		//IL_027c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0283: Unknown result type (might be due to invalid IL or missing references)
		//IL_028d: Expected O, but got Unknown
		List<string> list = new List<string>();
		XmlNodeList xmlNodeList = NC.A.SettingsXml.DocumentElement.SelectNodes(XC.A(3624));
		XmlDocument xmlDocument = new XmlDocument();
		new List<System.Windows.Controls.ListViewItem>();
		if (xmlDoc == null)
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
			xmlDoc = NC.A.SettingsXml;
		}
		xmlDocument.LoadXml(M.DefaultShortcuts);
		AllowedKeys = Base.ComboBoxKeys();
		AllItems = new ObservableCollection<ShortcutItem>();
		AdornerPadding = new Thickness(0.0);
		this.m_A = true;
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
				AdornerLayer.GetAdornerLayer(this.m_A).Remove((Adorner)(object)this.m_A);
				this.m_A = null;
				this.m_A = null;
			}
		}
		XmlNodeList xmlNodeList2 = xmlDocument.DocumentElement.SelectNodes(XC.A(3661));
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = xmlNodeList2.GetEnumerator();
			while (enumerator.MoveNext())
			{
				XmlNode xmlNode = (XmlNode)enumerator.Current;
				try
				{
					string value = xmlNode.Attributes[XC.A(3678)].Value;
					if (!Shortcuts.Dictionary.TryGetValue(value, out var value2))
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
							switch (7)
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
								switch (4)
								{
								case 0:
									continue;
								}
								string text = value2.FriendlyName;
								if (text.StartsWith(XC.A(3689)))
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
									text = xmlNodeList.Item(checked(Conversions.ToInteger(Regex.Replace(text, XC.A(3714), "")) - 1)).Attributes[XC.A(3725)].Value;
								}
								AllItems.Add(new ShortcutItem(value2.Category, value, text, value2.Utility, xmlDoc.SelectSingleNode(Shortcuts.ShortcutXpath(value)).Attributes[XC.A(3734)].Value, value2.DefaultKey, value2.Description));
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
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (5)
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
		SourceCollection.GroupDescriptions.Add(new PropertyGroupDescription(XC.A(3753)));
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
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Expected O, but got Unknown
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
			System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
			if (comboBox.IsLoaded)
			{
				ShortcutItem val = (ShortcutItem)comboBox.DataContext;
				if (val.LastKey.Length > 0)
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
					A(val);
				}
				val = null;
			}
			comboBox = null;
			return;
		}
	}

	private void ModifyKeystroke(object sender, System.Windows.Input.KeyEventArgs e)
	{
		//IL_0042: Unknown result type (might be due to invalid IL or missing references)
		//IL_0048: Expected O, but got Unknown
		if (lvShortcuts.SelectedItems.Count <= 0)
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
					if (!flag)
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
						modifierShift = (flag2 ? 1 : 0);
					}
					else
					{
						modifierShift = 1;
					}
				}
				else
				{
					modifierShift = 0;
				}
				obj.ModifierShift = (byte)modifierShift != 0;
				_ = null;
				if (flag)
				{
					goto IL_0108;
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
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
					if (flag3)
					{
						goto IL_0108;
					}
				}
				goto IL_017f;
			}
			goto IL_0186;
			IL_0186:
			val = null;
			return;
			IL_017f:
			this.m_A = false;
			goto IL_0186;
			IL_0108:
			Key key = e.Key;
			if ((uint)(key - 116) > 5u)
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
				if (key != Key.System)
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
					string keyFromPress = Base.GetKeyFromPress(e);
					if (!AllowedKeys.Contains(keyFromPress))
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
						AllowedKeys.Add(keyFromPress);
					}
					val.LastKey = keyFromPress;
					A(val);
					e.Handled = true;
				}
			}
			goto IL_017f;
		}
	}

	private void CtrlShift0(object sender, System.Windows.Input.KeyEventArgs e)
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_0022: Expected O, but got Unknown
		if (!Base.IsCtrlShift0(e, (ShortcutItem)lvShortcuts.SelectedItems[0]))
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
			D();
			return;
		}
	}

	private void A(ShortcutItem A)
	{
		if (this.A(A, B: false))
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
			if (!A.ModifierCtrl && !A.ModifierAlt)
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
				if (!A.ModifierShift)
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
					goto IL_004a;
				}
			}
			string text = Base.BuildKeystroke(A);
			if (Operators.CompareString(text, XC.A(3807), TextCompare: false) == 0)
			{
				while (true)
				{
					switch (5)
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
					switch (1)
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
						switch (7)
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
						if (Operators.CompareString(current.Keystroke, text, TextCompare: false) == 0 && current != A)
						{
							if (Base.ConfirmOverwrite(current))
							{
								current.IsEnabled = false;
								C(current);
								break;
							}
							return false;
						}
						continue;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_011d;
						}
						continue;
						end_IL_011d:
						break;
					}
					break;
				}
			}
			return true;
		}
		goto IL_004a;
		IL_004a:
		if (!B)
		{
			Forms.WarningMessage(XC.A(3770));
		}
		return false;
	}

	private bool A(string A)
	{
		string text = A.ToLower();
		uint num = VC.A(text);
		if (num <= 698900162)
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
			if (num <= 421688663)
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
				if (num <= 363347782)
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
					if (num <= 203579616)
					{
						if (num != 104072603)
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
							if (num != 203579616)
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
							else if (Operators.CompareString(text, XC.A(4451), TextCompare: false) == 0)
							{
								goto IL_0b21;
							}
						}
						else
						{
							if (Operators.CompareString(text, XC.A(4337), TextCompare: false) == 0)
							{
								goto IL_0b1d;
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
					else if (num != 329792544)
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
						if (num != 337800568)
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
							if (num != 363347782)
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
							else if (Operators.CompareString(text, XC.A(3845), TextCompare: false) == 0)
							{
								goto IL_0b09;
							}
						}
						else
						{
							if (Operators.CompareString(text, XC.A(4409), TextCompare: false) == 0)
							{
								goto IL_0b21;
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
					else
					{
						if (Operators.CompareString(text, XC.A(4014), TextCompare: false) == 0)
						{
							goto IL_0b11;
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
				else if (num <= 396903020)
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
					if (num != 380125401)
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
						if (num != 388133425)
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
							if (num != 396903020)
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
							}
							else
							{
								if (Operators.CompareString(text, XC.A(3962), TextCompare: false) == 0)
								{
									goto IL_0b11;
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
							if (Operators.CompareString(text, XC.A(4414), TextCompare: false) == 0)
							{
								goto IL_0b21;
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
					else if (Operators.CompareString(text, XC.A(3858), TextCompare: false) == 0)
					{
						goto IL_0b09;
					}
				}
				else if (num != 404911044)
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
					if (num != 413680639)
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
						if (num != 421688663)
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
						else if (Operators.CompareString(text, XC.A(4419), TextCompare: false) == 0)
						{
							goto IL_0b21;
						}
					}
					else if (Operators.CompareString(text, XC.A(4040), TextCompare: false) == 0)
					{
						goto IL_0b15;
					}
				}
				else
				{
					if (Operators.CompareString(text, XC.A(4441), TextCompare: false) == 0)
					{
						goto IL_0b21;
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
			else if (num <= 547901591)
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
				if (num <= 447235877)
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
					if (num != 438466282)
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
						if (num != 447235877)
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
						else if (Operators.CompareString(text, XC.A(4066), TextCompare: false) == 0)
						{
							goto IL_0b15;
						}
					}
					else
					{
						if (Operators.CompareString(text, XC.A(4446), TextCompare: false) == 0)
						{
							goto IL_0b21;
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
				else if (num != 497568734)
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
					if (num != 514346353)
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
						if (num != 547901591)
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
						else if (Operators.CompareString(text, XC.A(3884), TextCompare: false) == 0)
						{
							goto IL_0b0d;
						}
					}
					else
					{
						if (Operators.CompareString(text, XC.A(3975), TextCompare: false) == 0)
						{
							goto IL_0b11;
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
				else if (Operators.CompareString(text, XC.A(3949), TextCompare: false) == 0)
				{
					goto IL_0b11;
				}
			}
			else if (num <= 581456829)
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
				if (num != 564679210)
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
					if (num != 581318465)
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
						if (num == 581456829)
						{
							if (Operators.CompareString(text, XC.A(3897), TextCompare: false) == 0)
							{
								goto IL_0b0d;
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
						if (Operators.CompareString(text, XC.A(4320), TextCompare: false) == 0)
						{
							goto IL_0b1d;
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
				else if (Operators.CompareString(text, XC.A(4027), TextCompare: false) == 0)
				{
					goto IL_0b15;
				}
			}
			else if (num != 665344924)
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
				if (num != 682122543)
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
					if (num != 698900162)
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
						if (Operators.CompareString(text, XC.A(4079), TextCompare: false) == 0)
						{
							goto IL_0b15;
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
					if (Operators.CompareString(text, XC.A(4092), TextCompare: false) == 0)
					{
						goto IL_0b15;
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
			else
			{
				if (Operators.CompareString(text, XC.A(3923), TextCompare: false) == 0)
				{
					goto IL_0b0d;
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
		else if (num <= 1477700177)
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
			if (num <= 799565876)
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
				if (num <= 749233019)
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
					if (num != 732455400)
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
						if (num != 749233019)
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
							if (Operators.CompareString(text, XC.A(4053), TextCompare: false) == 0)
							{
								goto IL_0b15;
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
						if (Operators.CompareString(text, XC.A(4001), TextCompare: false) == 0)
						{
							goto IL_0b11;
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
				else if (num != 766010638)
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
					if (num != 782788257)
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
						if (num == 799565876)
						{
							if (Operators.CompareString(text, XC.A(3936), TextCompare: false) == 0)
							{
								goto IL_0b11;
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
					else
					{
						if (Operators.CompareString(text, XC.A(3871), TextCompare: false) == 0)
						{
							goto IL_0b09;
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
					if (Operators.CompareString(text, XC.A(3910), TextCompare: false) == 0)
					{
						goto IL_0b0d;
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
			else if (num <= 872303894)
			{
				if (num != 826951197)
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
					if (num != 849898733)
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
							if (Operators.CompareString(text, XC.A(4158), TextCompare: false) == 0)
							{
								goto IL_0b19;
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
					else if (Operators.CompareString(text, XC.A(3832), TextCompare: false) == 0)
					{
						goto IL_0b09;
					}
				}
				else
				{
					if (Operators.CompareString(text, XC.A(4139), TextCompare: false) == 0)
					{
						goto IL_0b19;
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
			else if (num != 1464394840)
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
				if (num != 1466661412)
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
					if (num == 1477700177)
					{
						if (Operators.CompareString(text, XC.A(4124), TextCompare: false) == 0)
						{
							goto IL_0b19;
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
				else
				{
					if (Operators.CompareString(text, XC.A(4179), TextCompare: false) == 0)
					{
						goto IL_0b19;
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
			else if (Operators.CompareString(text, XC.A(4105), TextCompare: false) == 0)
			{
				goto IL_0b19;
			}
		}
		else if (num <= 3160966811u)
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
			if (num <= 2355079881u)
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
				if (num != 1566262289)
				{
					if (num != 2355079881u)
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
						if (Operators.CompareString(text, XC.A(4237), TextCompare: false) == 0)
						{
							goto IL_0b19;
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
					if (Operators.CompareString(text, XC.A(4424), TextCompare: false) == 0)
					{
						goto IL_0b21;
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
					if (num != 3160966811u)
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
					}
					else
					{
						if (Operators.CompareString(text, XC.A(4301), TextCompare: false) == 0)
						{
							goto IL_0b1d;
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
					if (Operators.CompareString(text, XC.A(4375), TextCompare: false) == 0)
					{
						goto IL_0b1d;
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
			else if (Operators.CompareString(text, XC.A(4210), TextCompare: false) == 0)
			{
				goto IL_0b19;
			}
		}
		else if (num <= 3998789804u)
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
			if (num != 3308457076u)
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
				if (num != 3351539004u)
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
					if (num == 3998789804u)
					{
						if (Operators.CompareString(text, XC.A(4356), TextCompare: false) == 0)
						{
							goto IL_0b1d;
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
					if (Operators.CompareString(text, XC.A(4456), TextCompare: false) == 0)
					{
						goto IL_0b21;
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
				if (Operators.CompareString(text, XC.A(4392), TextCompare: false) == 0)
				{
					goto IL_0b1d;
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
		else if (num != 4037543175u)
		{
			if (num != 4081400426u)
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
				if (num != 4231138174u)
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
				else if (Operators.CompareString(text, XC.A(4471), TextCompare: false) == 0)
				{
					goto IL_0b21;
				}
			}
			else
			{
				if (Operators.CompareString(text, XC.A(4268), TextCompare: false) == 0)
				{
					goto IL_0b19;
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
		else if (Operators.CompareString(text, XC.A(3988), TextCompare: false) == 0)
		{
			goto IL_0b11;
		}
		return false;
		IL_0b11:
		return true;
		IL_0b19:
		return true;
		IL_0b21:
		return true;
		IL_0b0d:
		return true;
		IL_0b15:
		return true;
		IL_0b1d:
		return true;
		IL_0b09:
		return true;
	}

	private void B(ShortcutItem A)
	{
		Microsoft.Office.Interop.Word.Application application = PC.A.Application;
		object objectValue = RuntimeHelpers.GetObjectValue(application.CustomizationContext);
		Template template = null;
		string text = Base.BuildKeystroke(A);
		try
		{
			template = Shortcuts.GetMacabacusTemplate(application);
			XmlDocument settingsXml = NC.A.SettingsXml;
			XmlNode xmlNode = this.A(settingsXml, A);
			string value = xmlNode.Attributes[XC.A(3734)].Value;
			int num = Shortcuts.ConvertKeystroke(text, application);
			try
			{
				KeyBindings keyBindings = application.KeyBindings;
				object KeyCode = RuntimeHelpers.GetObjectValue(Missing.Value);
				keyBindings.Key(num, ref KeyCode).Clear();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			Shortcuts.AssignShortcut((WdKey)num, xmlNode.Attributes[XC.A(3678)].Value, application);
			xmlNode.Attributes[XC.A(3734)].Value = text;
			if (value.Length > 0)
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
				Microsoft.Office.Interop.Word.Application application2 = application;
				int keyCode = Shortcuts.ConvertKeystroke(value, application);
				object KeyCode = RuntimeHelpers.GetObjectValue(Missing.Value);
				((_Application)application2).get_FindKey(keyCode, ref KeyCode).Clear();
			}
			xmlNode = null;
			this.A(settingsXml);
			A.Keystroke = text;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.ErrorMessage(XC.A(4478) + text + XC.A(4616) + ex4.Message);
			ProjectData.ClearProjectError();
		}
		finally
		{
			XmlDocument settingsXml = null;
		}
		if (template != null)
		{
			template.Saved = true;
			template = null;
			application.CustomizationContext = RuntimeHelpers.GetObjectValue(objectValue);
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
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			NC.A.SaveSettings(A);
			Shortcuts.dictLookup2 = null;
			NC.A.Invalidate();
			return;
		}
	}

	private void SelectOnMouseDown(object sender, MouseButtonEventArgs e)
	{
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0011: Expected O, but got Unknown
		ShortcutItem val = (ShortcutItem)((System.Windows.Controls.Control)sender).DataContext;
		if (!val.IsSelected)
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
			val.IsSelected = true;
			D();
			e.Handled = true;
		}
		val = null;
	}

	private void ShortcutSelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		//IL_005f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0069: Expected O, but got Unknown
		if (e.OriginalSource != lvShortcuts)
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
			if (e.RemovedItems.Count > 0)
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
				ShortcutItem A = (ShortcutItem)e.RemovedItems[0];
				if (A.IsDirty && A.IsEnabled)
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
					if (!this.A(A, B: true))
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
						if (System.Windows.Forms.MessageBox.Show(XC.A(4629) + A.FriendlyName + XC.A(4726), XC.A(2438), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
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
			if (lvShortcuts.SelectedItems.Count <= 0)
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
				btnReset.IsEnabled = this.m_B;
				return;
			}
		}
	}

	private void ShortcutCheckedChanged(object sender, RoutedEventArgs e)
	{
		//IL_003a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0040: Expected O, but got Unknown
		if (!this.m_B)
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
			if (this.m_A)
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
				ShortcutItem val = (ShortcutItem)((System.Windows.Controls.CheckBox)sender).DataContext;
				this.m_A = true;
				val.Clear();
				if (!val.IsEnabled)
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
					C(val);
				}
				this.m_A = false;
				val = null;
				return;
			}
		}
	}

	private void C(ShortcutItem A)
	{
		Microsoft.Office.Interop.Word.Application application = PC.A.Application;
		XmlDocument settingsXml = NC.A.SettingsXml;
		XmlNode xmlNode = this.A(settingsXml, A);
		XmlAttribute xmlAttribute = xmlNode.Attributes[XC.A(3734)];
		if (xmlAttribute.Value.Length > 0)
		{
			Microsoft.Office.Interop.Word.Application application2 = application;
			int keyCode = Shortcuts.ConvertKeystroke(xmlNode.Attributes[XC.A(3734)].Value, application);
			object KeyCode = RuntimeHelpers.GetObjectValue(Missing.Value);
			((_Application)application2).get_FindKey(keyCode, ref KeyCode).Clear();
		}
		xmlAttribute.Value = "";
		_ = null;
		this.A(settingsXml);
		settingsXml = null;
		xmlNode = null;
		application = null;
	}

	private void SpacebarToggle(object sender, System.Windows.Input.KeyEventArgs e)
	{
		Base.SpacebarToggle(lvShortcuts, e);
	}

	private void ResetShortcut(object sender, RoutedEventArgs e)
	{
		//IL_002f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0035: Expected O, but got Unknown
		bool flag = false;
		if (lvShortcuts.SelectedItems.Count <= 0)
		{
			return;
		}
		ShortcutItem val = (ShortcutItem)lvShortcuts.SelectedItems[0];
		IEnumerator<ShortcutItem> enumerator = default(IEnumerator<ShortcutItem>);
		try
		{
			enumerator = AllItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				ShortcutItem current = enumerator.Current;
				if (Operators.CompareString(current.Keystroke, val.DefaultKeystroke, TextCompare: false) != 0 || current == val)
				{
					continue;
				}
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
					Forms.WarningMessage(XC.A(4749) + current.FriendlyName + XC.A(4860));
					flag = true;
					break;
				}
				break;
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
		if (!flag)
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
			this.m_A = true;
			val.IsEnabled = true;
			this.m_A = false;
			val.Reset();
			B(val);
		}
		val = null;
	}

	private void ResetAll(object sender, RoutedEventArgs e)
	{
		if (!Base.ConfirmResetAll())
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
			XmlDocument settingsXml = NC.A.SettingsXml;
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(M.DefaultShortcuts);
			XmlNode xmlNode = settingsXml.DocumentElement.SelectSingleNode(XC.A(4863));
			xmlNode.ParentNode.RemoveChild(xmlNode);
			settingsXml.DocumentElement.SelectSingleNode(XC.A(4910)).AppendChild(settingsXml.ImportNode(xmlDocument.DocumentElement, deep: true));
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
		if (System.Windows.Forms.MessageBox.Show(XC.A(4929), XC.A(2438), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
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
			Microsoft.Office.Interop.Word.Application application = PC.A.Application;
			this.m_A = true;
			XmlDocument settingsXml;
			try
			{
				settingsXml = NC.A.SettingsXml;
				using (IEnumerator<ShortcutItem> enumerator = AllItems.GetEnumerator())
				{
					while (enumerator.MoveNext())
					{
						ShortcutItem current = enumerator.Current;
						ShortcutItem val = current;
						if (val.IsEnabled)
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
							val.IsEnabled = false;
							XmlNode xmlNode = A(settingsXml, current);
							Microsoft.Office.Interop.Word.Application application2 = application;
							int keyCode = Shortcuts.ConvertKeystroke(xmlNode.Attributes[XC.A(3734)].Value, application);
							object KeyCode = RuntimeHelpers.GetObjectValue(Missing.Value);
							((_Application)application2).get_FindKey(keyCode, ref KeyCode).Clear();
							xmlNode.Attributes[XC.A(3734)].Value = "";
							xmlNode = null;
						}
						val = null;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_0127;
						}
						continue;
						end_IL_0127:
						break;
					}
				}
				A(settingsXml);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(ex2.Message);
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
		//IL_0096: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a0: Expected O, but got Unknown
		GridViewColumnHeader gridViewColumnHeader = (GridViewColumnHeader)sender;
		string text = gridViewColumnHeader.Tag.ToString();
		if (this.m_A != null)
		{
			AdornerLayer.GetAdornerLayer(this.m_A).Remove((Adorner)(object)this.m_A);
			SourceCollection.SortDescriptions.Clear();
		}
		SourceCollection.GroupDescriptions.Clear();
		ListSortDirection listSortDirection = ListSortDirection.Descending;
		if (this.m_A == gridViewColumnHeader && this.m_A.Direction == listSortDirection)
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
			listSortDirection = ListSortDirection.Ascending;
		}
		this.m_A = gridViewColumnHeader;
		this.m_A = new SortAdorner((UIElement)this.m_A, listSortDirection);
		AdornerLayer.GetAdornerLayer(this.m_A).Add((Adorner)(object)this.m_A);
		SourceCollection.SortDescriptions.Add(new SortDescription(text, listSortDirection));
		if (Operators.CompareString(text, XC.A(5251), TextCompare: false) == 0)
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
		if (e.WidthChanged)
		{
			UpdateColumnWidths(lvShortcuts, e);
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
		if (!(this.m_A.VerticalOffset > 0.0))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			base.Dispatcher.BeginInvoke((global::A.A)([SpecialName] () =>
			{
				Base.UpdateColumnWidth(gvShortcuts.Columns[1]);
			}), DispatcherPriority.ContextIdle);
			return;
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

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (!this.m_C)
		{
			this.m_C = true;
			Uri resourceLocator = new Uri(XC.A(5266), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
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
				switch (7)
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
				switch (4)
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
				switch (2)
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
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					((GridViewColumnHeader)target).Click += SortColumn;
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
					((GridViewColumnHeader)target).Click += SortColumn;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 12:
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				btnReset = (System.Windows.Controls.Button)target;
				return;
			}
		case 13:
			btnDefaults = (System.Windows.Controls.Button)target;
			break;
		case 14:
			btnClear = (System.Windows.Controls.Button)target;
			break;
		default:
			this.m_C = true;
			break;
		}
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
		if (connectionId == 5)
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
			((System.Windows.Controls.CheckBox)target).Checked += ShortcutCheckedChanged;
			((System.Windows.Controls.CheckBox)target).Unchecked += ShortcutCheckedChanged;
			((System.Windows.Controls.CheckBox)target).PreviewMouseDown += SelectOnMouseDown;
		}
		if (connectionId == 7)
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
			((System.Windows.Controls.CheckBox)target).PreviewMouseDown += SelectOnMouseDown;
		}
		if (connectionId == 8)
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
			((System.Windows.Controls.CheckBox)target).PreviewMouseDown += SelectOnMouseDown;
		}
		if (connectionId == 9)
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
			((System.Windows.Controls.CheckBox)target).PreviewMouseDown += SelectOnMouseDown;
		}
		if (connectionId != 10)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += LastKeyChanged;
			((System.Windows.Controls.ComboBox)target).PreviewMouseDown += SelectOnMouseDown;
			return;
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
