using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Threading;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.Config;
using MacabacusMacros.Config.Settings;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.Check;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Check;
using PowerPointAddIn1.DeckCheck.Errors;
using PowerPointAddIn1.DeckCheck.Reformat;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.DeckCheck.UI;

[DesignerGenerated]
public sealed class wpfPane : System.Windows.Controls.UserControl, INotifyPropertyChanged, IComponentConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static ThreadStart A;

		public static Func<BaseError, int> A;

		public static Func<BaseError, float> A;

		public static Func<BaseError, float> B;

		public static Func<BaseError, float> C;

		public static Func<BaseError, float> D;

		public static Func<BaseError, int> B;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal void A()
		{
		}

		[SpecialName]
		internal int A(BaseError A)
		{
			if (A.Slide != null)
			{
				return A.Slide.SlideIndex;
			}
			return 0;
		}

		[SpecialName]
		internal float A(BaseError A)
		{
			if (A.Shape != null)
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
						return A.Shape.Left;
					}
				}
			}
			float result = default(float);
			return result;
		}

		[SpecialName]
		internal float B(BaseError A)
		{
			if (A.Shape != null)
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
						return A.Shape.Top;
					}
				}
			}
			float result = default(float);
			return result;
		}

		[SpecialName]
		internal float C(BaseError A)
		{
			if (((BaseError)A).TextRanges != null)
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
						return ((BaseError)A).TextRanges[0].BoundTop;
					}
				}
			}
			float result = default(float);
			return result;
		}

		[SpecialName]
		internal float D(BaseError A)
		{
			if (((BaseError)A).TextRanges != null)
			{
				return ((BaseError)A).TextRanges[0].BoundLeft;
			}
			float result = default(float);
			return result;
		}

		[SpecialName]
		internal int B(BaseError A)
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0006: Unknown result type (might be due to invalid IL or missing references)
			//IL_0009: Expected I4, but got Unknown
			return (int)((BaseError)A).Severity;
		}
	}

	[CompilerGenerated]
	internal sealed class BD
	{
		public string A;

		public wpfPane A;

		[SpecialName]
		internal void A()
		{
			this.A.txtLoading.Text = this.A;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private static FieldInfo m_A;

	private BackgroundWorker m_A;

	private List<string> m_A;

	private wpfWarnings m_A;

	private wpfReformat m_A;

	private wpfSettings m_A;

	private Visibility m_A;

	[CompilerGenerated]
	private bool m_A;

	private bool? m_A;

	[AccessedThroughProperty("popRefresh")]
	[CompilerGenerated]
	private Popup m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnRefreshAll")]
	private System.Windows.Controls.Button m_A;

	[AccessedThroughProperty("btnRefreshSelection")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	[AccessedThroughProperty("grdHeader")]
	[CompilerGenerated]
	private Grid m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnRefresh")]
	private System.Windows.Controls.Button m_C;

	[AccessedThroughProperty("radErrorsView")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_A;

	[AccessedThroughProperty("radInfoView")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_B;

	[AccessedThroughProperty("txtSearch")]
	[CompilerGenerated]
	private System.Windows.Controls.TextBox m_A;

	[AccessedThroughProperty("stkWarnings")]
	[CompilerGenerated]
	private StackPanel m_A;

	[AccessedThroughProperty("chkErrors")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkWarnings")]
	private System.Windows.Controls.CheckBox m_B;

	[AccessedThroughProperty("chkMessages")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnSettings")]
	private System.Windows.Controls.Button m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("grdLoading")]
	private Grid m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("iconWorking")]
	private StackPanel m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("txtLoading")]
	private TextBlock m_A;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_E;

	[AccessedThroughProperty("iconSuccess")]
	[CompilerGenerated]
	private StackPanel m_C;

	[AccessedThroughProperty("grdContent")]
	[CompilerGenerated]
	private Grid m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("conWarnings")]
	private ContentControl m_A;

	[AccessedThroughProperty("conReformat")]
	[CompilerGenerated]
	private ContentControl m_B;

	[AccessedThroughProperty("conSettings")]
	[CompilerGenerated]
	private ContentControl m_C;

	private bool m_B;

	public wpfWarnings WarningsView
	{
		get
		{
			return this.m_A;
		}
		private set
		{
			this.m_A = value;
			A(AH.A(57536));
		}
	}

	public wpfReformat ReformatView
	{
		get
		{
			return this.m_A;
		}
		private set
		{
			this.m_A = value;
			A(AH.A(57561));
		}
	}

	public wpfSettings SettingsView
	{
		get
		{
			return this.m_A;
		}
		private set
		{
			this.m_A = value;
			A(AH.A(57586));
		}
	}

	public Visibility WarningsViewVisibility
	{
		get
		{
			return this.m_A;
		}
		private set
		{
			this.m_A = value;
			A(AH.A(57611));
		}
	}

	public bool WholePresentation
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

	internal virtual Popup popRefresh
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
			System.Windows.Input.KeyEventHandler value2 = ClosePopup;
			EventHandler value3 = RefreshPopupOpened;
			EventHandler value4 = RefreshPopupClosed;
			Popup popup = this.m_A;
			if (popup != null)
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
				popup.PreviewKeyDown -= value2;
				popup.Opened -= value3;
				popup.Closed -= value4;
			}
			this.m_A = value;
			popup = this.m_A;
			if (popup == null)
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
				popup.PreviewKeyDown += value2;
				popup.Opened += value3;
				popup.Closed += value4;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnRefreshAll
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
			RoutedEventHandler value2 = RefreshAll;
			System.Windows.Controls.Button button = this.m_A;
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
			this.m_A = value;
			button = this.m_A;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnRefreshSelection
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
			RoutedEventHandler value2 = RefreshSelection;
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

	internal virtual Grid grdHeader
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

	internal virtual System.Windows.Controls.Button btnRefresh
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
			RoutedEventHandler value2 = btnRefresh_Click;
			System.Windows.Controls.Button button = this.m_C;
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.RadioButton radErrorsView
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

	internal virtual System.Windows.Controls.RadioButton radInfoView
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

	internal virtual System.Windows.Controls.TextBox txtSearch
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
			KeyboardFocusChangedEventHandler value2 = SearchBoxGotFocus;
			KeyboardFocusChangedEventHandler value3 = SearchBoxLostFocus;
			TextChangedEventHandler value4 = SearchTextChanged;
			RoutedEventHandler value5 = SearchTextBoxLostFocus;
			RoutedEventHandler value6 = SearchTextBoxGotFocus;
			System.Windows.Input.KeyEventHandler value7 = SearchTextBoxLostFocus;
			System.Windows.Controls.TextBox textBox = this.m_A;
			if (textBox != null)
			{
				textBox.GotKeyboardFocus -= value2;
				textBox.LostKeyboardFocus -= value3;
				textBox.TextChanged -= value4;
				textBox.LostFocus -= value5;
				textBox.GotFocus -= value6;
				textBox.KeyDown -= value7;
			}
			this.m_A = value;
			textBox = this.m_A;
			if (textBox == null)
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
				textBox.GotKeyboardFocus += value2;
				textBox.LostKeyboardFocus += value3;
				textBox.TextChanged += value4;
				textBox.LostFocus += value5;
				textBox.GotFocus += value6;
				textBox.KeyDown += value7;
				return;
			}
		}
	}

	internal virtual StackPanel stkWarnings
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

	internal virtual System.Windows.Controls.CheckBox chkErrors
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

	internal virtual System.Windows.Controls.CheckBox chkWarnings
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

	internal virtual System.Windows.Controls.CheckBox chkMessages
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
			this.m_C = value;
		}
	}

	internal virtual System.Windows.Controls.Button btnSettings
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = ShowSettingsView;
			System.Windows.Controls.Button button = this.m_D;
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
			this.m_D = value;
			button = this.m_D;
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

	internal virtual Grid grdLoading
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

	internal virtual StackPanel iconWorking
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

	internal virtual TextBlock txtLoading
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

	internal virtual System.Windows.Controls.Button btnCancel
	{
		[CompilerGenerated]
		get
		{
			return this.m_E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = CancelButtonClicked;
			System.Windows.Controls.Button button = this.m_E;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.m_E = value;
			button = this.m_E;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual StackPanel iconSuccess
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
			this.m_C = value;
		}
	}

	internal virtual Grid grdContent
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
			this.m_C = value;
		}
	}

	internal virtual ContentControl conWarnings
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

	internal virtual ContentControl conReformat
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

	internal virtual ContentControl conSettings
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
			this.m_C = value;
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
				switch (6)
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
		base.Loaded += wpfPane_Loaded;
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
		this.m_A = Visibility.Visible;
		InitializeComponent();
		Pane.TaskPane = this;
		A();
		if (!SharedSettings.IsSettingEditable(Constants.XML_PROOFING_RULES))
		{
			btnSettings.IsEnabled = false;
			btnSettings.ToolTip = AH.A(58133);
		}
		this.m_A = new BackgroundWorker();
		BackgroundWorker a = this.m_A;
		a.WorkerSupportsCancellation = true;
		a.WorkerReportsProgress = false;
		a.DoWork += StartCheckBackground;
		a.RunWorkerCompleted += bgw_RunWorkerCompleted;
		_ = null;
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

	private static void A()
	{
		wpfPane.m_A = typeof(SystemParameters).GetField(AH.A(57499), BindingFlags.Static | BindingFlags.NonPublic);
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
			switch (3)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if ((object)wpfPane.m_A == null)
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
				wpfPane.m_A.SetValue(null, false);
				return;
			}
		}
	}

	private void wpfPane_Loaded(object sender, RoutedEventArgs e)
	{
		chkErrors.IsChecked = PB.Settings.ProofingShowErrors;
		chkWarnings.IsChecked = PB.Settings.ProofingShowWarnings;
		chkMessages.IsChecked = PB.Settings.ProofingShowMessages;
		chkErrors.Checked += ToggleErrors;
		chkErrors.Unchecked += ToggleErrors;
		chkWarnings.Checked += ToggleWarnings;
		chkWarnings.Unchecked += ToggleWarnings;
		chkMessages.Checked += ToggleMessages;
		chkMessages.Unchecked += ToggleMessages;
		radErrorsView.Checked += [SpecialName] (object a0, RoutedEventArgs a1) =>
		{
			ShowWarningsView();
		};
		radInfoView.Checked += ShowReformatView;
	}

	private void ClosePopup(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Escape)
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
			popRefresh.IsOpen = false;
			e.Handled = true;
			return;
		}
	}

	private void RefreshPopupOpened(object sender, EventArgs e)
	{
		btnRefresh.IsEnabled = false;
	}

	private void RefreshPopupClosed(object sender, EventArgs e)
	{
		btnRefresh.IsEnabled = true;
	}

	private void btnRefresh_Click(object sender, RoutedEventArgs e)
	{
		popRefresh.IsOpen = true;
		btnRefreshAll.Focus();
	}

	private void RefreshAll(object sender, RoutedEventArgs e)
	{
		A(A: true);
	}

	private void RefreshSelection(object sender, RoutedEventArgs e)
	{
		A(A: false);
	}

	private void A(bool A)
	{
		popRefresh.IsOpen = false;
		WholePresentation = A;
		C();
	}

	private void C()
	{
		Pane.ActiveItem = null;
		WarningsView = null;
		ReformatView = null;
		ShowLoadingState();
		StartCheckAsynch();
	}

	private void B(string A)
	{
		try
		{
			base.Dispatcher.Invoke([SpecialName] () =>
			{
				txtLoading.Text = A;
			});
			Dispatcher dispatcher = base.Dispatcher;
			ThreadStart method;
			if (_Closure_0024__.A == null)
			{
				method = (_Closure_0024__.A = [SpecialName] () =>
				{
				});
			}
			else
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
				method = _Closure_0024__.A;
			}
			dispatcher.Invoke(DispatcherPriority.Background, method);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			Analysis analysis = Main.Analysis;
			object obj;
			if (analysis == null)
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
				obj = null;
			}
			else
			{
				obj = analysis.Errors;
			}
			if (obj == null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						Forms.WarningMessage(AH.A(57954));
						return;
					}
				}
			}
			int count = Main.Analysis.Errors.Count;
			this.m_A = count > 0;
			if (this.m_A == true)
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
				grdLoading.Visibility = Visibility.Collapsed;
			}
			else
			{
				iconWorking.Visibility = Visibility.Hidden;
				iconSuccess.Visibility = Visibility.Visible;
			}
			grdContent.Visibility = Visibility.Visible;
			if (radErrorsView.IsChecked == true)
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
				WarningsView = new wpfWarnings(Main.Analysis.Errors);
				if (count > 0)
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
					conWarnings.Visibility = Visibility.Visible;
					try
					{
						System.Windows.Forms.Application.DoEvents();
						WarningsView.ScrollSpy(NG.A.Application.ActiveWindow.Selection.SlideRange[1]);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				else
				{
					conWarnings.Visibility = Visibility.Collapsed;
				}
			}
			else if (radInfoView.IsChecked == true)
			{
				ReformatView = new wpfReformat(Main.Analysis.Conventions);
				B(A: true);
			}
			List<string> a = this.m_A;
			if (a == null)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
			if (a.Any())
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						this.m_A.Insert(0, AH.A(58059));
						Forms.WarningMessage(string.Join(string.Format(AH.A(58120), AH.A(7894), AH.A(7894)), this.m_A));
						return;
					}
				}
			}
		});
		Pane.CloseCallout();
	}

	public void ShowLoadingState()
	{
		iconSuccess.Visibility = Visibility.Hidden;
		iconWorking.Visibility = Visibility.Visible;
		grdLoading.Visibility = Visibility.Visible;
		grdContent.Visibility = Visibility.Collapsed;
		System.Windows.Forms.Application.DoEvents();
	}

	public void StartCheckAsynch()
	{
		this.m_A.RunWorkerAsync();
	}

	private void StartCheckBackground(object sender, DoWorkEventArgs e)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			this.m_A = new List<string>();
			try
			{
				D();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				this.m_A.Add(string.Format(AH.A(47123), AH.A(7894), ex2.Message));
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
		});
	}

	private void D()
	{
		//IL_00de: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e3: Unknown result type (might be due to invalid IL or missing references)
		//IL_0100: Unknown result type (might be due to invalid IL or missing references)
		//IL_0117: Unknown result type (might be due to invalid IL or missing references)
		//IL_011c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0238: Unknown result type (might be due to invalid IL or missing references)
		//IL_023d: Unknown result type (might be due to invalid IL or missing references)
		//IL_026b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0270: Unknown result type (might be due to invalid IL or missing references)
		//IL_0295: Unknown result type (might be due to invalid IL or missing references)
		//IL_029a: Unknown result type (might be due to invalid IL or missing references)
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = NG.A.Application.ActivePresentation;
		SlideRange slideRange = null;
		if (WholePresentation)
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
			slideRange = activePresentation.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		else
		{
			try
			{
				slideRange = NG.A.Application.ActiveWindow.Selection.SlideRange;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				activePresentation = null;
				ProjectData.ClearProjectError();
				return;
			}
		}
		B(AH.A(57656));
		Main.Analysis = new Analysis(activePresentation, this.m_A);
		B(AH.A(57707));
		Analysis analysis = Main.Analysis;
		if (WholePresentation)
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
			if ((int)analysis.Options.TemplateRules != 0)
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
				TemplateRules.Check(activePresentation);
			}
			if ((int)analysis.Options.MultipleSlideMasters != 0)
			{
				Miscellaneous.CheckSlideMasterCount(activePresentation, slideRange);
			}
			if ((int)analysis.Options.SlideCount != 0)
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
				Miscellaneous.CheckSlideCount(activePresentation, analysis.Options.MaxSlides);
			}
		}
		IEnumerator enumerator2 = default(IEnumerator);
		foreach (Slide item in slideRange)
		{
			try
			{
				B(AH.A(57756) + item.SlideIndex);
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
				continue;
			}
			Checks checks = analysis.Checks;
			if (checks.SlideTitleCapitalization != null)
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
				checks.SlideTitleCapitalization.Check(item);
			}
			if (checks.SlideNumbers != null)
			{
				checks.SlideNumbers.Check(item);
			}
			if (checks.HiddenSlides != null)
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
				checks.HiddenSlides.Check(item);
			}
			if (checks.MisalignedShapes != null)
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
				checks.MisalignedShapes.Check(item);
			}
			checks = null;
			if ((int)analysis.Options.SlideWordCount != 0)
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
				Miscellaneous.CheckSlideWordCount(item, analysis.Options.MaxSlideWords);
			}
			if ((int)analysis.Options.Hyperlinks != 0)
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
				new PowerPointAddIn1.DeckCheck.Check.Hyperlinks().Check(item);
				_ = null;
			}
			_ = analysis.Options.CheckAgendaUpdated;
			List<int> D = new List<int>();
			item.Application.StartNewUndoEntry();
			bool flag = false;
			Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = item.Shapes.Range(RuntimeHelpers.GetObjectValue(Missing.Value));
			try
			{
				enumerator2 = shapeRange.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shape a = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
					if (!A(a, item, Main.Analysis, ref D))
					{
						continue;
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
					flag = true;
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_0319;
					}
					continue;
					end_IL_0319:
					break;
				}
			}
			finally
			{
				if (enumerator2 is IDisposable)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						(enumerator2 as IDisposable).Dispose();
						break;
					}
				}
			}
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
				item.Application.CommandBars.ExecuteMso(AH.A(40491));
				System.Windows.Forms.Application.DoEvents();
			}
			if (analysis.Checks.FootnoteReferenceMissing != null)
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
				analysis.Checks.FootnoteReferenceMissing.Check(item, analysis.Conventions.FootnoteNumbers[item], D, analysis.Conventions.FootnoteShapes[item]);
			}
			D = null;
			shapeRange = null;
		}
		if (analysis.Errors.Count > 0)
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
			try
			{
				Analysis analysis2 = analysis;
				List<BaseError> errors = analysis.Errors;
				Func<BaseError, int> keySelector;
				if (_Closure_0024__.A == null)
				{
					keySelector = (_Closure_0024__.A = [SpecialName] (BaseError A) => (A.Slide != null) ? A.Slide.SlideIndex : 0);
				}
				else
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
					keySelector = _Closure_0024__.A;
				}
				IOrderedEnumerable<BaseError> source = errors.OrderBy(keySelector);
				Func<BaseError, float> keySelector2;
				if (_Closure_0024__.A == null)
				{
					keySelector2 = (_Closure_0024__.A = [SpecialName] (BaseError A) =>
					{
						if (A.Shape != null)
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
									return A.Shape.Left;
								}
							}
						}
						float result = default(float);
						return result;
					});
				}
				else
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
					keySelector2 = _Closure_0024__.A;
				}
				IOrderedEnumerable<BaseError> source2 = source.ThenBy(keySelector2).ThenBy([SpecialName] (BaseError A) =>
				{
					if (A.Shape != null)
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
								return A.Shape.Top;
							}
						}
					}
					float result = default(float);
					return result;
				});
				Func<BaseError, float> keySelector3;
				if (_Closure_0024__.C == null)
				{
					keySelector3 = (_Closure_0024__.C = [SpecialName] (BaseError A) =>
					{
						if (((BaseError)A).TextRanges != null)
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
									return ((BaseError)A).TextRanges[0].BoundTop;
								}
							}
						}
						float result = default(float);
						return result;
					});
				}
				else
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
					keySelector3 = _Closure_0024__.C;
				}
				IOrderedEnumerable<BaseError> source3 = source2.ThenBy(keySelector3);
				Func<BaseError, float> keySelector4;
				if (_Closure_0024__.D == null)
				{
					float num = default(float);
					keySelector4 = (_Closure_0024__.D = [SpecialName] (BaseError A) => (((BaseError)A).TextRanges != null) ? ((BaseError)A).TextRanges[0].BoundLeft : num);
				}
				else
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
					keySelector4 = _Closure_0024__.D;
				}
				IOrderedEnumerable<BaseError> source4 = source3.ThenBy(keySelector4);
				Func<BaseError, int> keySelector5;
				if (_Closure_0024__.B == null)
				{
					keySelector5 = (_Closure_0024__.B = [SpecialName] (BaseError A) => (int)((BaseError)A).Severity);
				}
				else
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
					keySelector5 = _Closure_0024__.B;
				}
				analysis2.Errors = source4.ThenBy(keySelector5).ToList();
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			B(AH.A(57787));
		}
		analysis = null;
		activePresentation = null;
	}

	private bool A(Microsoft.Office.Interop.PowerPoint.Shape A, Slide B, Analysis C, ref List<int> D)
	{
		//IL_0a3a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a3f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0204: Unknown result type (might be due to invalid IL or missing references)
		//IL_0209: Unknown result type (might be due to invalid IL or missing references)
		//IL_03ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_03b0: Unknown result type (might be due to invalid IL or missing references)
		//IL_03cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_03d2: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a0: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_0321: Unknown result type (might be due to invalid IL or missing references)
		//IL_0331: Unknown result type (might be due to invalid IL or missing references)
		//IL_0343: Unknown result type (might be due to invalid IL or missing references)
		//IL_0355: Unknown result type (might be due to invalid IL or missing references)
		//IL_0367: Unknown result type (might be due to invalid IL or missing references)
		//IL_0379: Unknown result type (might be due to invalid IL or missing references)
		//IL_0389: Unknown result type (might be due to invalid IL or missing references)
		//IL_0399: Unknown result type (might be due to invalid IL or missing references)
		//IL_02e2: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f6: Unknown result type (might be due to invalid IL or missing references)
		//IL_030a: Unknown result type (might be due to invalid IL or missing references)
		//IL_04b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_04b6: Unknown result type (might be due to invalid IL or missing references)
		//IL_051b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0520: Unknown result type (might be due to invalid IL or missing references)
		//IL_04c6: Unknown result type (might be due to invalid IL or missing references)
		//IL_04cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0591: Unknown result type (might be due to invalid IL or missing references)
		//IL_04fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_0502: Unknown result type (might be due to invalid IL or missing references)
		//IL_0506: Unknown result type (might be due to invalid IL or missing references)
		//IL_05bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_05c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_0607: Unknown result type (might be due to invalid IL or missing references)
		//IL_060c: Unknown result type (might be due to invalid IL or missing references)
		//IL_08b4: Unknown result type (might be due to invalid IL or missing references)
		//IL_08b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0916: Unknown result type (might be due to invalid IL or missing references)
		//IL_091b: Unknown result type (might be due to invalid IL or missing references)
		//IL_08d6: Unknown result type (might be due to invalid IL or missing references)
		//IL_08db: Unknown result type (might be due to invalid IL or missing references)
		bool result = false;
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		if (shape.Type != MsoShapeType.msoGroup)
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
			if (shape.Visible == MsoTriState.msoTrue)
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
				if (C.Checks.CheckPlaceholders)
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
					if (A.Type == MsoShapeType.msoPlaceholder)
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
						bool flag = false;
						using (List<Microsoft.Office.Interop.PowerPoint.Shape>.Enumerator enumerator = C.Conventions.LayoutPlaceholders[B].GetEnumerator())
						{
							while (true)
							{
								if (enumerator.MoveNext())
								{
									Microsoft.Office.Interop.PowerPoint.Shape current = enumerator.Current;
									if (current.PlaceholderFormat.Type != A.PlaceholderFormat.Type || !Helpers.IsShapeMatch(current, A))
									{
										continue;
									}
									flag = true;
									Checks checks = C.Checks;
									if (checks.PlaceholderFillMismatch != null)
									{
										checks.PlaceholderFillMismatch.Check(B, A, current);
									}
									if (checks.PlaceholderFontStyleMismatch != null)
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
										checks.PlaceholderFontStyleMismatch.Check(B, A, current);
									}
									if (checks.PlaceholderFontColorMismatch != null)
									{
										checks.PlaceholderFontColorMismatch.Check(B, A, current);
									}
									if (checks.PlaceholderBulletMismatch != null)
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
										checks.PlaceholderBulletMismatch.Check(B, A, current);
									}
									if (checks.PlaceholderIndentMismatch != null)
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
										checks.PlaceholderIndentMismatch.Check(B, A, current);
									}
									if (checks.PlaceholderMarginsMismatch != null)
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
										checks.PlaceholderMarginsMismatch.Check(B, A, current);
									}
									checks = null;
									break;
								}
								while (true)
								{
									switch (4)
									{
									case 0:
										break;
									default:
										goto end_IL_01bb;
									}
									continue;
									end_IL_01bb:
									break;
								}
								break;
							}
						}
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
							Main.Analysis.Errors.Add(new PlaceholderLayoutMismatch(B, A));
						}
					}
				}
				if ((int)((Settings)C.Options).ColorPalette != 0)
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
					MsoShapeType type = A.Type;
					if (type <= MsoShapeType.msoPlaceholder)
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
						if (type != MsoShapeType.msoAutoShape)
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
							if (type != MsoShapeType.msoChart)
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
								if (type != MsoShapeType.msoPlaceholder)
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
									goto IL_03a3;
								}
							}
						}
					}
					else if (type != MsoShapeType.msoTextBox)
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
						if (type != MsoShapeType.msoTable)
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
							if (type != MsoShapeType.msoSmartArt)
							{
								goto IL_03a3;
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
					Severity colorPalette = ((Settings)C.Options).ColorPalette;
					Checks checks2 = C.Checks;
					if (A.HasChart == MsoTriState.msoFalse)
					{
						bool blnCheckPlaceholderFillColor = checks2.PlaceholderFillMismatch == null;
						bool blnCheckPlaceholderFontColor = checks2.PlaceholderFontColorMismatch == null;
						ShapeColors.FillColor(B, A, ((Checks)checks2).PaletteColors, colorPalette, blnCheckPlaceholderFillColor);
						ShapeColors.TextColor(B, A, ((Checks)checks2).PaletteColors, colorPalette, blnCheckPlaceholderFontColor);
						ShapeColors.BorderColor(B, A, ((Checks)checks2).PaletteColors, colorPalette);
					}
					else
					{
						ChartColors.ChartArea(B, A, ((Checks)checks2).PaletteColors, colorPalette);
						ChartColors.PlotArea(B, A, ((Checks)checks2).PaletteColors, colorPalette);
						ChartColors.Series(B, A, ((Checks)checks2).PaletteColors, colorPalette);
						ChartColors.Axes(B, A, ((Checks)checks2).PaletteColors, colorPalette);
						ChartColors.Title(B, A, ((Checks)checks2).PaletteColors, colorPalette);
						ChartColors.Legend(B, A, ((Checks)checks2).PaletteColors, colorPalette);
						ChartColors.DataTable(B, A, ((Checks)checks2).PaletteColors, colorPalette);
						ChartColors.Groups(B, A, ((Checks)checks2).PaletteColors, colorPalette);
					}
					checks2 = null;
				}
				goto IL_03a3;
			}
			this.A(C, B, A);
		}
		else if (shape.Visible == MsoTriState.msoTrue)
		{
			int count = shape.GroupItems.Count;
			for (int i = 1; i <= count; i = checked(i + 1))
			{
				this.A(shape.GroupItems[i], B, C, ref D);
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
		else
		{
			this.A(C, B, A);
		}
		goto IL_0a30;
		IL_03a3:
		if ((int)((Settings)C.Options).FillTransparency != 0)
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
			ShapeColors.FillTransparency(B, A);
		}
		if ((int)((Settings)C.Options).FillGradients != 0)
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
			ShapeColors.FillGradient(B, A);
		}
		Checks checks3 = C.Checks;
		if (checks3.RotatedShapes != null)
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
			checks3.RotatedShapes.Check(B, A);
		}
		if (checks3.ShapeOutOfBounds != null)
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
			checks3.ShapeOutOfBounds.Check(B, A);
		}
		if (checks3.ShapeOutsideMargins != null)
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
			checks3.ShapeOutsideMargins.Check(B, A);
		}
		if (checks3.OverlappingText != null)
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
			checks3.OverlappingText.Check(B, A);
		}
		if (checks3.MasterShapePosition != null)
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
			checks3.MasterShapePosition.Check(B, A);
		}
		checks3 = null;
		Settings options = C.Options;
		if ((int)options.ShapeEffects == 0)
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
			if ((int)options.TextEffects == 0)
			{
				goto IL_0510;
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
		if (A.Type != MsoShapeType.msoInkComment && A.Type != MsoShapeType.msoInk)
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
			Effects.AllEffects(B, A, options.ShapeEffects, options.TextEffects);
		}
		goto IL_0510;
		IL_0589:
		if ((int)((Settings)C.Options).CheckLinks != 0)
		{
			Miscellaneous.CheckMacabacusLink(B, A);
		}
		if (A.HasChart == MsoTriState.msoTrue)
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
			if ((int)C.Options.MinMaxFontSize != 0)
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
				ChartFonts.A(B, A, Main.Analysis.Conventions.MinFontSize, Main.Analysis.Conventions.MaxFontSize);
			}
			if ((int)((Settings)C.Options).CheckChartElements != 0)
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
				ChartElements.Legend(B, A);
				ChartElements.MissingDataLabels(B, A);
				ChartElements.DataLabelConsistency(B, A);
				ChartElements.DataLabelFormat(B, A);
			}
		}
		else if (A.HasTable == MsoTriState.msoTrue)
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
			List<CellMargins> list = new List<CellMargins>();
			List<Microsoft.Office.Interop.PowerPoint.Shape> list2 = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
			Table table = A.Table;
			int count2 = table.Rows.Count;
			int count3 = table.Columns.Count;
			int num = count2;
			checked
			{
				for (int j = 1; j <= num; j++)
				{
					int num2 = count3;
					for (int k = 1; k <= num2; k++)
					{
						if (table.Cell(j, k).Shape.HasTextFrame != MsoTriState.msoTrue)
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
							break;
						}
						Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = table.Cell(j, k).Shape.TextFrame2;
						try
						{
							list.Add(new CellMargins
							{
								Top = textFrame.MarginTop,
								Right = textFrame.MarginRight,
								Bottom = textFrame.MarginBottom,
								Left = textFrame.MarginLeft
							});
							list2.Add(A.Table.Cell(j, k).Shape);
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						if (textFrame.HasText == MsoTriState.msoTrue)
						{
							this.A(C, B, A, A.Table.Cell(j, k).Shape.TextFrame2.TextRange, ref D);
						}
						textFrame = null;
					}
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
				table = null;
				Tables.CellMargins(B, A, list, list2);
				list = null;
				list2 = null;
			}
		}
		else if (A.HasSmartArt == MsoTriState.msoTrue)
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
			try
			{
				IEnumerator enumerator2 = A.SmartArt.AllNodes.GetEnumerator();
				try
				{
					while (enumerator2.MoveNext())
					{
						SmartArtNode smartArtNode = (SmartArtNode)enumerator2.Current;
						if (smartArtNode.TextFrame2.HasText == MsoTriState.msoTrue)
						{
							this.A(C, B, A, smartArtNode.TextFrame2.TextRange, ref D);
						}
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0861;
						}
						continue;
						end_IL_0861:
						break;
					}
				}
				finally
				{
					IDisposable disposable = enumerator2 as IDisposable;
					if (disposable != null)
					{
						disposable.Dispose();
					}
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		else if (PowerPointAddIn1.Shapes.Images.HasPictureOrGraphic(A))
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
			if ((int)((Settings)C.Options).ImageDistortion != 0)
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
				PowerPointAddIn1.DeckCheck.Check.Images.Distortion(B, A);
				result = true;
			}
			if ((int)((Settings)C.Options).ImageCropping != 0)
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
				PowerPointAddIn1.DeckCheck.Check.Images.Cropping(B, A);
			}
			PowerPointAddIn1.DeckCheck.Check.Images.AirplaneMode(B, A);
		}
		else if (PowerPointAddIn1.Shapes.Images.HasPicture(A))
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
			if ((int)((Settings)C.Options).LinkedPictures != 0)
			{
				PowerPointAddIn1.DeckCheck.Check.Images.LinkedPicture(B, A);
			}
		}
		else
		{
			MsoShapeType type2 = A.Type;
			if (type2 != MsoShapeType.msoAutoShape)
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
				if (type2 != MsoShapeType.msoPlaceholder)
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
					if (type2 != MsoShapeType.msoTextBox)
					{
						goto IL_0a30;
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
			if (A.HasTextFrame == MsoTriState.msoTrue)
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
				PowerPointAddIn1.DeckCheck.Check.Fonts.A(B, A);
				if (A.TextFrame2.HasText == MsoTriState.msoTrue)
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
					this.A(C, B, A, A.TextFrame2.TextRange, ref D);
				}
			}
		}
		goto IL_0a30;
		IL_0510:
		options = null;
		if ((int)((Settings)C.Options).Ink != 0)
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
			if (A.Type != MsoShapeType.msoInkComment)
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
				if (A.Type != MsoShapeType.msoInk)
				{
					goto IL_0589;
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
			if (A.Visible == MsoTriState.msoTrue)
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
				Main.Analysis.Errors.Add(new Ink(B, A));
			}
		}
		goto IL_0589;
		IL_0a30:
		shape = null;
		if ((int)C.Options.Animation != 0)
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
			bool flag2 = false;
			try
			{
				if (A.ParentGroup != null)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						flag2 = true;
						break;
					}
				}
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			if (!flag2)
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
				Effects.Animation(B, A);
			}
		}
		return result;
	}

	private void A(Analysis A, Slide B, Microsoft.Office.Interop.PowerPoint.Shape C)
	{
		Checks checks = A.Checks;
		if (checks.HiddenShapes != null)
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
			checks.HiddenShapes.Check(B, C);
		}
		checks = null;
	}

	private void A(Analysis A, Slide B, Microsoft.Office.Interop.PowerPoint.Shape C, TextRange2 D, ref List<int> E)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Unknown result type (might be due to invalid IL or missing references)
		//IL_003f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0291: Unknown result type (might be due to invalid IL or missing references)
		//IL_0054: Unknown result type (might be due to invalid IL or missing references)
		//IL_0059: Unknown result type (might be due to invalid IL or missing references)
		//IL_0078: Unknown result type (might be due to invalid IL or missing references)
		//IL_007d: Unknown result type (might be due to invalid IL or missing references)
		Checks checks = A.Checks;
		if ((int)A.Options.MinMaxFontSize != 0)
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
			PowerPointAddIn1.DeckCheck.Check.Fonts.A(B, C, D);
		}
		if ((int)A.Options.FractionalFontSize != 0)
		{
			PowerPointAddIn1.DeckCheck.Check.Fonts.B(B, C, D);
		}
		if ((int)((Settings)A.Options).StrikethroughFont != 0)
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
			PowerPointAddIn1.DeckCheck.Check.Fonts.C(B, C, D);
		}
		if ((int)A.Options.IllegalFonts != 0)
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
			PowerPointAddIn1.DeckCheck.Check.Fonts.A(B, C, D, A.Conventions.LegalFonts);
		}
		foreach (BaseTextCheck shapeTextCheck in checks.ShapeTextChecks)
		{
			shapeTextCheck.Check(B, C, D, D.Text);
		}
		int count = D.get_Paragraphs(-1, -1).Count;
		for (int i = 1; i <= count; i = checked(i + 1))
		{
			TextRange2 textRange = D.get_Paragraphs(i, -1);
			string strText = Text.PrintableText(textRange.Text);
			using (List<BaseTextCheck>.Enumerator enumerator2 = checks.ParagraphTextChecks.GetEnumerator())
			{
				while (enumerator2.MoveNext())
				{
					enumerator2.Current.Check(B, C, textRange, strText);
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_014d;
					}
					continue;
					end_IL_014d:
					break;
				}
			}
			if (checks.FootnoteExplanationMissing != null)
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
				checks.FootnoteExplanationMissing.Check(B, C, textRange, ref E);
			}
			if (checks.MultipleFontFamilies != null)
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
				checks.MultipleFontFamilies.Check(B, C, textRange);
			}
			if (checks.LineSpacing != null)
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
				checks.LineSpacing.CheckParagraph(textRange);
			}
			if (textRange.ParagraphFormat.Bullet.Type != MsoBulletType.msoBulletNone)
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
				try
				{
					if (checks.BulletPunctuation != null)
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
						checks.BulletPunctuation.CheckParagraph(textRange, strText);
					}
					if (checks.BulletSize != null)
					{
						checks.BulletSize.CheckParagraph(textRange);
					}
					if (checks.BulletIndentation != null)
					{
						checks.BulletIndentation.CheckParagraph(textRange);
					}
					if (D.ParagraphFormat.Bullet.Type != MsoBulletType.msoBulletPicture)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							if (checks.BulletFontFamily != null)
							{
								checks.BulletFontFamily.CheckParagraph(textRange, strText);
							}
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
				if ((int)A.Options.BulletWordCount != 0)
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
					Miscellaneous.CheckBulletWordCount(B, C, textRange, A.Options.MaxBulletWords);
				}
			}
			if (checks.ProofingLanguage == null)
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
			checks.ProofingLanguage.CheckParagraph(textRange);
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			if (checks.LineSpacing != null)
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
				checks.LineSpacing.Check(B, C);
			}
			if (checks.BulletPunctuation != null)
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
				checks.BulletPunctuation.Check(B, C);
			}
			if (checks.BulletSize != null)
			{
				checks.BulletSize.Check(B, C);
			}
			if (checks.BulletIndentation != null)
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
				checks.BulletIndentation.Check(B, C);
			}
			if (checks.BulletFontFamily != null)
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
				checks.BulletFontFamily.Check(B, C);
			}
			if (checks.ProofingLanguage != null)
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
				checks.ProofingLanguage.Check(B, C);
			}
			checks = null;
			return;
		}
	}

	private void CancelButtonClicked(object sender, RoutedEventArgs e)
	{
		E();
	}

	private void E()
	{
		if (this.m_A == null)
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
			if (!this.m_A.IsBusy)
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
				this.m_A.CancelAsync();
				return;
			}
		}
	}

	public void ScrollSpy(Slide sld)
	{
		if (WarningsView != null && WarningsViewVisibility == Visibility.Visible)
		{
			WarningsView.ScrollSpy(sld);
		}
	}

	public void ShowWarningsView()
	{
		if (WarningsView == null)
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
			if (SettingsView != null)
			{
				if (SettingsView.SettingsDirty)
				{
					goto IL_0055;
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
			WarningsView = new wpfWarnings(Main.Analysis.Errors);
		}
		goto IL_0055;
		IL_0055:
		B(A: false);
		conSettings.Visibility = Visibility.Collapsed;
		conWarnings.Visibility = Visibility.Visible;
		WarningsViewVisibility = Visibility.Visible;
		txtSearch.Visibility = Visibility.Visible;
		stkWarnings.Visibility = Visibility.Visible;
		if (SettingsView == null)
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
			if (!SettingsView.SettingsDirty)
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
				C();
				SettingsView.SettingsDirty = false;
				return;
			}
		}
	}

	private void ShowReformatView(object sender, RoutedEventArgs e)
	{
		if (!A())
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
					radErrorsView.IsChecked = true;
					return;
				}
			}
		}
		if (ReformatView == null)
		{
			ReformatView = new wpfReformat(Main.Analysis.Conventions);
		}
		conWarnings.Visibility = Visibility.Collapsed;
		conSettings.Visibility = Visibility.Collapsed;
		B(A: true);
		WarningsViewVisibility = Visibility.Hidden;
		txtSearch.Visibility = Visibility.Collapsed;
		stkWarnings.Visibility = Visibility.Collapsed;
	}

	private void B(bool A)
	{
		conReformat.Visibility = ((!A) ? Visibility.Collapsed : Visibility.Visible);
		Grid grid = grdLoading;
		int visibility;
		if (object.Equals(this.m_A, false))
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
			if (!A)
			{
				visibility = 0;
				goto IL_0058;
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
		visibility = 2;
		goto IL_0058;
		IL_0058:
		grid.Visibility = (Visibility)visibility;
	}

	private void ShowSettingsView(object sender, RoutedEventArgs e)
	{
		if (!A())
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
			if (SettingsView == null)
			{
				SettingsView = new wpfSettings(this);
			}
			conWarnings.Visibility = Visibility.Collapsed;
			B(A: false);
			conSettings.Visibility = Visibility.Visible;
			WarningsViewVisibility = Visibility.Hidden;
			grdHeader.Visibility = Visibility.Collapsed;
			grdLoading.Visibility = Visibility.Collapsed;
			return;
		}
	}

	internal void F()
	{
		grdHeader.Visibility = Visibility.Visible;
		ShowWarningsView();
	}

	private bool A()
	{
		return Access.AllowPowerPointOperation((PlanType)4, (Restriction)2, false);
	}

	private void SearchBoxGotFocus(object sender, KeyboardFocusChangedEventArgs e)
	{
		System.Windows.Controls.TextBox obj = (System.Windows.Controls.TextBox)sender;
		object obj2 = ColorConverter.ConvertFromString(AH.A(57824));
		obj.Background = new SolidColorBrush((obj2 != null) ? ((Color)obj2) : default(Color));
		obj.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
		obj.Padding = new Thickness(24.0, 0.0, 13.0, 0.0);
		_ = null;
	}

	private void SearchBoxLostFocus(object sender, KeyboardFocusChangedEventArgs e)
	{
		System.Windows.Controls.TextBox textBox = (System.Windows.Controls.TextBox)sender;
		if (textBox.Text.Length == 0)
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
			textBox.Background = new SolidColorBrush(System.Windows.Media.Colors.White);
			textBox.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
			textBox.Padding = new Thickness(0.0);
		}
		textBox = null;
	}

	private void SearchTextChanged(object sender, TextChangedEventArgs e)
	{
		if (txtSearch.Text.Length != 0)
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
			H();
			return;
		}
	}

	private void SearchTextBoxLostFocus(object sender, RoutedEventArgs e)
	{
		if (txtSearch.Text.Length <= 0)
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
			H();
			return;
		}
	}

	private void SearchTextBoxGotFocus(object sender, RoutedEventArgs e)
	{
		if (txtSearch.Text.Length <= 0)
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
			txtSearch.SelectAll();
			return;
		}
	}

	private void SearchTextBoxLostFocus(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key == Key.Return && txtSearch.Text.Length > 0)
		{
			txtSearch.SelectAll();
			H();
		}
	}

	private void ToggleErrors(object sender, RoutedEventArgs e)
	{
		G();
		PB.Settings.ProofingShowErrors = chkErrors.IsChecked.Value;
	}

	private void ToggleWarnings(object sender, RoutedEventArgs e)
	{
		G();
		PB.Settings.ProofingShowWarnings = chkWarnings.IsChecked.Value;
	}

	private void ToggleMessages(object sender, RoutedEventArgs e)
	{
		G();
		PB.Settings.ProofingShowMessages = chkMessages.IsChecked.Value;
	}

	private void G()
	{
		WarningsView.ToggleItems();
	}

	private void H()
	{
		WarningsView.ApplyFilter();
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (this.m_B)
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
			this.m_B = true;
			Uri resourceLocator = new Uri(AH.A(57839), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
			return;
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
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					popRefresh = (Popup)target;
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
					btnRefreshAll = (System.Windows.Controls.Button)target;
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
					btnRefreshSelection = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			grdHeader = (Grid)target;
			return;
		}
		if (connectionId == 5)
		{
			btnRefresh = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					radErrorsView = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
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
					radInfoView = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					txtSearch = (System.Windows.Controls.TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			stkWarnings = (StackPanel)target;
			return;
		}
		if (connectionId == 10)
		{
			chkErrors = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkWarnings = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			chkMessages = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnSettings = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					grdLoading = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 15)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					iconWorking = (StackPanel)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtLoading = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnCancel = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					iconSuccess = (StackPanel)target;
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			grdContent = (Grid)target;
			return;
		}
		if (connectionId == 20)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					conWarnings = (ContentControl)target;
					return;
				}
			}
		}
		if (connectionId == 21)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					conReformat = (ContentControl)target;
					return;
				}
			}
		}
		if (connectionId == 22)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					conSettings = (ContentControl)target;
					return;
				}
			}
		}
		this.m_B = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}

	[SpecialName]
	[CompilerGenerated]
	private void I()
	{
		Analysis analysis = Main.Analysis;
		object obj;
		if (analysis == null)
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
			obj = null;
		}
		else
		{
			obj = analysis.Errors;
		}
		if (obj == null)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					Forms.WarningMessage(AH.A(57954));
					return;
				}
			}
		}
		int count = Main.Analysis.Errors.Count;
		this.m_A = count > 0;
		if (this.m_A == true)
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
			grdLoading.Visibility = Visibility.Collapsed;
		}
		else
		{
			iconWorking.Visibility = Visibility.Hidden;
			iconSuccess.Visibility = Visibility.Visible;
		}
		grdContent.Visibility = Visibility.Visible;
		if (radErrorsView.IsChecked == true)
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
			WarningsView = new wpfWarnings(Main.Analysis.Errors);
			if (count > 0)
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
				conWarnings.Visibility = Visibility.Visible;
				try
				{
					System.Windows.Forms.Application.DoEvents();
					WarningsView.ScrollSpy(NG.A.Application.ActiveWindow.Selection.SlideRange[1]);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			else
			{
				conWarnings.Visibility = Visibility.Collapsed;
			}
		}
		else if (radInfoView.IsChecked == true)
		{
			ReformatView = new wpfReformat(Main.Analysis.Conventions);
			B(A: true);
		}
		List<string> a = this.m_A;
		if (a == null)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		if (!a.Any())
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
			this.m_A.Insert(0, AH.A(58059));
			Forms.WarningMessage(string.Join(string.Format(AH.A(58120), AH.A(7894), AH.A(7894)), this.m_A));
			return;
		}
	}

	[SpecialName]
	[CompilerGenerated]
	private void J()
	{
		this.m_A = new List<string>();
		try
		{
			D();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			this.m_A.Add(string.Format(AH.A(47123), AH.A(7894), ex2.Message));
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}
}
