using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Sheets;

[DesignerGenerated]
public sealed class wpfSheetActivate : System.Windows.Window, IComponentConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<KeyValuePair<string, object>, string> A;

		public static Func<KeyValuePair<string, object>, object> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal string A(KeyValuePair<string, object> A)
		{
			return A.Key;
		}

		[SpecialName]
		internal object A(KeyValuePair<string, object> A)
		{
			return A.Value;
		}
	}

	[CompilerGenerated]
	internal sealed class DF
	{
		public Microsoft.Office.Interop.Excel.Workbook A;

		public DF(DF A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(KeyValuePair<string, object> A)
		{
			return Operators.ConditionalCompareObjectEqual(A.Key, NewLateBinding.LateGet(this.A.ActiveSheet, null, VH.A(19019), new object[0], null, null, null), TextCompare: false);
		}
	}

	private Dictionary<string, object> m_A;

	private readonly string m_A;

	[AccessedThroughProperty("txtSearch")]
	[CompilerGenerated]
	private TextBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lbxSheets")]
	private System.Windows.Controls.ListBox m_A;

	[AccessedThroughProperty("btnOk")]
	[CompilerGenerated]
	private Button m_A;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private Button B;

	private bool m_A;

	internal virtual TextBox txtSearch
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
			RoutedEventHandler value2 = txtSearch_GotFocus;
			RoutedEventHandler value3 = txtSearch_LostFocus;
			TextBox textBox = this.m_A;
			if (textBox != null)
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
				textBox.GotFocus -= value2;
				textBox.LostFocus -= value3;
			}
			this.m_A = value;
			textBox = this.m_A;
			if (textBox == null)
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
				textBox.GotFocus += value2;
				textBox.LostFocus += value3;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.ListBox lbxSheets
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
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnOk_Click;
			Button button = this.m_A;
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
			this.m_A = value;
			button = this.m_A;
			if (button != null)
			{
				button.Click += value2;
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

	public wpfSheetActivate()
	{
		base.Loaded += wpfSheetActivate_Loaded;
		base.Closing += wpfSheetActivate_Closing;
		base.PreviewKeyDown += DialogPreviewKeyDown;
		this.m_A = VH.A(102604);
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void wpfSheetActivate_Loaded(object sender, RoutedEventArgs e)
	{
		DF a = default(DF);
		DF CS_0024_003C_003E8__locals6 = new DF(a);
		CS_0024_003C_003E8__locals6.A = null;
		try
		{
			CS_0024_003C_003E8__locals6.A = MH.A.Application.ActiveWorkbook;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (CS_0024_003C_003E8__locals6.A == null)
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
					Forms.WarningMessage(System.Windows.Window.GetWindow(this), VH.A(83055));
					Close();
					return;
				}
			}
		}
		this.m_A = new Dictionary<string, object>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = CS_0024_003C_003E8__locals6.A.Sheets.GetEnumerator();
			while (enumerator.MoveNext())
			{
				object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
				this.m_A.Add(Conversions.ToString(NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null)), RuntimeHelpers.GetObjectValue(objectValue));
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
		lbxSheets.ItemsSource = this.m_A;
		lbxSheets.SelectedItem = this.m_A.Where([SpecialName] (KeyValuePair<string, object> A) => Operators.ConditionalCompareObjectEqual(A.Key, NewLateBinding.LateGet(CS_0024_003C_003E8__locals6.A.ActiveSheet, null, VH.A(19019), new object[0], null, null, null), TextCompare: false)).ElementAtOrDefault(0);
		A();
		CS_0024_003C_003E8__locals6.A = null;
		txtSearch.TextChanged += txtSearch_TextChanged;
	}

	private void A()
	{
		System.Windows.Controls.ListBox listBox = lbxSheets;
		listBox.UpdateLayout();
		listBox.ScrollIntoView(RuntimeHelpers.GetObjectValue(listBox.Items.GetItemAt(listBox.SelectedIndex)));
		((ListBoxItem)listBox.ItemContainerGenerator.ContainerFromItem(RuntimeHelpers.GetObjectValue(listBox.SelectedItem))).Focus();
		listBox = null;
	}

	private void wpfSheetActivate_Closing(object sender, CancelEventArgs e)
	{
		if (base.DialogResult == true)
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
			try
			{
				if (lbxSheets.SelectedIndex > -1)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						NewLateBinding.LateCall(((Dictionary<string, object>)lbxSheets.ItemsSource).ElementAt(lbxSheets.SelectedIndex).Value, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
						break;
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		this.m_A = null;
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = true;
	}

	private void txtSearch_GotFocus(object sender, RoutedEventArgs e)
	{
		TextBox textBox = txtSearch;
		if (Operators.CompareString(textBox.Text, this.m_A, TextCompare: false) == 0)
		{
			textBox.TextChanged -= txtSearch_TextChanged;
			textBox.Text = "";
			textBox.TextChanged += txtSearch_TextChanged;
		}
		else if (textBox.Text.Length > 0)
		{
			textBox.SelectAll();
		}
		textBox.Foreground = new SolidColorBrush(SystemColors.WindowTextColor);
		textBox = null;
	}

	private void txtSearch_LostFocus(object sender, RoutedEventArgs e)
	{
		TextBox textBox = txtSearch;
		if (textBox.Text.Length == 0)
		{
			textBox.Foreground = new SolidColorBrush(SystemColors.ScrollBarColor);
			textBox.TextChanged -= txtSearch_TextChanged;
			textBox.Text = this.m_A;
			textBox.TextChanged += txtSearch_TextChanged;
		}
		else
		{
			textBox.Foreground = new SolidColorBrush(SystemColors.WindowTextColor);
		}
		textBox = null;
	}

	private void txtSearch_TextChanged(object sender, TextChangedEventArgs e)
	{
		Dictionary<string, object> itemsSource;
		if (txtSearch.Text.Length != 0)
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
			if (Operators.CompareString(txtSearch.Text, this.m_A, TextCompare: false) != 0)
			{
				IEnumerable<KeyValuePair<string, object>> source = this.m_A.Where([SpecialName] (KeyValuePair<string, object> A) => A.Key.ToLower().Contains(txtSearch.Text.ToLower()));
				Func<KeyValuePair<string, object>, string> keySelector = [SpecialName] (KeyValuePair<string, object> A) => A.Key;
				Func<KeyValuePair<string, object>, object> elementSelector;
				if (_Closure_0024__.A == null)
				{
					elementSelector = (_Closure_0024__.A = [SpecialName] (KeyValuePair<string, object> A) => A.Value);
				}
				else
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
					elementSelector = _Closure_0024__.A;
				}
				itemsSource = source.ToDictionary(keySelector, elementSelector);
				goto IL_00ce;
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
		itemsSource = this.m_A;
		goto IL_00ce;
		IL_00ce:
		System.Windows.Controls.ListBox listBox = lbxSheets;
		listBox.ItemsSource = itemsSource;
		if (listBox.Items.Count > 0)
		{
			listBox.SelectedIndex = 0;
			listBox.ScrollIntoView(RuntimeHelpers.GetObjectValue(listBox.Items.GetItemAt(0)));
		}
		listBox = null;
		itemsSource = null;
	}

	private void DialogPreviewKeyDown(object sender, KeyEventArgs e)
	{
		if (e.Key == Key.F)
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
			if (System.Windows.Input.Keyboard.Modifiers == ModifierKeys.Control)
			{
				txtSearch.Focus();
				e.Handled = true;
				return;
			}
		}
		if (txtSearch.IsKeyboardFocusWithin)
		{
			switch (e.Key)
			{
			case Key.Prior:
			case Key.Next:
			case Key.End:
			case Key.Home:
			case Key.Up:
			case Key.Down:
				lbxSheets.Focus();
				break;
			case Key.Left:
			case Key.Right:
				break;
			}
			return;
		}
		if (e.Key >= Key.A)
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
			if (e.Key <= Key.Z)
			{
				goto IL_015d;
			}
		}
		if (e.Key >= Key.D0)
		{
			if (e.Key <= Key.D9)
			{
				goto IL_015d;
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
		if (e.Key >= Key.NumPad0)
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
			if (e.Key <= Key.NumPad9)
			{
				goto IL_015d;
			}
		}
		if (e.Key != Key.OemPeriod)
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
			if (e.Key != Key.OemComma)
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
				if (e.Key != Key.Space && e.Key != Key.Delete)
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
					if (e.Key != Key.Back)
					{
						return;
					}
				}
			}
		}
		goto IL_015d;
		IL_015d:
		txtSearch.Focus();
		txtSearch.SelectionStart = txtSearch.Text.Length;
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_A)
		{
			this.m_A = true;
			Uri resourceLocator = new Uri(VH.A(102493), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
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
					txtSearch = (TextBox)target;
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
					lbxSheets = (System.Windows.Controls.ListBox)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (5)
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
				switch (1)
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

	[SpecialName]
	[CompilerGenerated]
	private bool A(KeyValuePair<string, object> A)
	{
		return A.Key.ToLower().Contains(txtSearch.Text.ToLower());
	}
}
