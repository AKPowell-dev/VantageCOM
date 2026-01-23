using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using A;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.Check;
using MacabacusMacros.Proofing.UI;
using MacabacusMacros.UI;
using Macabacus_Word.Proofing.Check;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.UI;

[DesignerGenerated]
public sealed class wpfPane : System.Windows.Controls.UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
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

		public static Func<BaseError, bool> A;

		public static Func<BaseError, bool> B;

		public static Func<BaseError, int> C;

		public static Func<BaseError, BaseError> A;

		public static Func<int, IEnumerable<BaseError>, F<int, IEnumerable<BaseError>>> A;

		public static Func<F<int, IEnumerable<BaseError>>, int> A;

		public static Func<F<int, IEnumerable<BaseError>>, G<int, int>> A;

		public static Func<BaseError, bool> C;

		public static Func<BaseError, bool> D;

		public static Func<BaseError, bool> E;

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
			return A.PageNumber;
		}

		[SpecialName]
		internal float A(BaseError A)
		{
			if (A.Ranges != null)
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
						return Conversions.ToSingle(A.Ranges[0].get_Information(WdInformation.wdVerticalPositionRelativeToPage));
					}
				}
			}
			float result = default(float);
			return result;
		}

		[SpecialName]
		internal float B(BaseError A)
		{
			if (A.Ranges != null)
			{
				return Conversions.ToSingle(A.Ranges[0].get_Information(WdInformation.wdHorizontalPositionRelativeToPage));
			}
			float result = default(float);
			return result;
		}

		[SpecialName]
		internal float C(BaseError A)
		{
			if (((BaseError)A).TextRanges != null)
			{
				return ((BaseError)A).TextRanges[0].BoundTop;
			}
			float result = default(float);
			return result;
		}

		[SpecialName]
		internal float D(BaseError A)
		{
			if (((BaseError)A).TextRanges != null)
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
						return ((BaseError)A).TextRanges[0].BoundLeft;
					}
				}
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

		[SpecialName]
		internal bool A(BaseError A)
		{
			return A.Range == null;
		}

		[SpecialName]
		internal bool B(BaseError A)
		{
			return A.Range != null;
		}

		[SpecialName]
		internal int C(BaseError A)
		{
			return A.PageNumber;
		}

		[SpecialName]
		internal BaseError A(BaseError A)
		{
			return A;
		}

		[SpecialName]
		internal F<int, IEnumerable<BaseError>> A(int A, IEnumerable<BaseError> B)
		{
			return new F<int, IEnumerable<BaseError>>(A, B);
		}

		[SpecialName]
		internal int A(F<int, IEnumerable<BaseError>> A)
		{
			return A.idx;
		}

		[SpecialName]
		internal G<int, int> A(F<int, IEnumerable<BaseError>> A)
		{
			return new G<int, int>(A.idx, A.Group.Count());
		}

		[SpecialName]
		internal bool C(BaseError A)
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0006: Unknown result type (might be due to invalid IL or missing references)
			//IL_0009: Invalid comparison between Unknown and I4
			return (int)((BaseError)A).Severity == 3;
		}

		[SpecialName]
		internal bool D(BaseError A)
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0006: Unknown result type (might be due to invalid IL or missing references)
			//IL_0009: Invalid comparison between Unknown and I4
			return (int)((BaseError)A).Severity == 2;
		}

		[SpecialName]
		internal bool E(BaseError A)
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0007: Invalid comparison between Unknown and I4
			return (int)((BaseError)A).Severity == 1;
		}
	}

	[CompilerGenerated]
	internal sealed class IC
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
	internal sealed class JC
	{
		public int A;

		public JC(JC A)
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
		internal bool A(BaseError A)
		{
			if (A.Range != null)
			{
				return A.PageNumber == this.A;
			}
			return false;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private BackgroundWorker m_A;

	private DispatcherTimer m_A;

	private readonly double m_A;

	private ICollectionView m_A;

	private static ObservableCollection<BaseError> m_A;

	private List<BaseError> m_A;

	private static int m_A = -1;

	private wpfReformat m_A;

	private wpfSettings m_A;

	private bool m_A;

	private bool m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("grdHeader")]
	private Grid m_A;

	[AccessedThroughProperty("btnRefresh")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkErrors")]
	private System.Windows.Controls.CheckBox m_A;

	[AccessedThroughProperty("chkWarnings")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_B;

	[AccessedThroughProperty("chkMessages")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_C;

	[AccessedThroughProperty("radErrorsView")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("radInfoView")]
	private System.Windows.Controls.RadioButton m_B;

	[AccessedThroughProperty("radSettingsView")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_C;

	[AccessedThroughProperty("txtSearch")]
	[CompilerGenerated]
	private System.Windows.Controls.TextBox m_A;

	[AccessedThroughProperty("stkLoading")]
	[CompilerGenerated]
	private StackPanel m_A;

	[AccessedThroughProperty("cvsLoading")]
	[CompilerGenerated]
	private Canvas m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("noFreeze")]
	private RotateTransform m_A;

	[AccessedThroughProperty("txtLoading")]
	[CompilerGenerated]
	private TextBlock m_A;

	[AccessedThroughProperty("conReformat")]
	[CompilerGenerated]
	private System.Windows.Controls.ContentControl m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("conSettings")]
	private System.Windows.Controls.ContentControl m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("lbxResults")]
	private System.Windows.Controls.ListBox m_A;

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

	private ObservableCollection<BaseError> AllItems
	{
		get
		{
			return wpfPane.m_A;
		}
		set
		{
			wpfPane.m_A = value;
		}
	}

	private List<BaseError> ItemsQueuedToRemove
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	private int NavigateAfterRemoveIndex
	{
		get
		{
			return wpfPane.m_A;
		}
		set
		{
			wpfPane.m_A = value;
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
			A(XC.A(38739));
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
			A(XC.A(38764));
		}
	}

	private bool CollapseAnimationRunning
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	private bool SettingsDirty
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
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
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnRefresh_Click;
			System.Windows.Controls.Button button = this.m_A;
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

	internal virtual System.Windows.Controls.RadioButton radSettingsView
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
			TextChangedEventHandler value2 = SearchTextChanged;
			RoutedEventHandler value3 = SearchTextBoxLostFocus;
			RoutedEventHandler value4 = SearchTextBoxGotFocus;
			System.Windows.Input.KeyEventHandler value5 = SearchTextBoxLostFocus;
			System.Windows.Controls.TextBox textBox = this.m_A;
			if (textBox != null)
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
				textBox.TextChanged -= value2;
				textBox.LostFocus -= value3;
				textBox.GotFocus -= value4;
				textBox.KeyDown -= value5;
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
				textBox.TextChanged += value2;
				textBox.LostFocus += value3;
				textBox.GotFocus += value4;
				textBox.KeyDown += value5;
				return;
			}
		}
	}

	internal virtual StackPanel stkLoading
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

	internal virtual Canvas cvsLoading
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

	internal virtual RotateTransform noFreeze
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

	internal virtual System.Windows.Controls.ContentControl conReformat
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

	internal virtual System.Windows.Controls.ContentControl conSettings
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

	internal virtual System.Windows.Controls.ListBox lbxResults
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
			System.Windows.Input.KeyEventHandler value2 = lbxResults_KeyDown;
			System.Windows.Controls.ListBox listBox = this.m_A;
			if (listBox != null)
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
				listBox.PreviewKeyDown -= value2;
			}
			this.m_A = value;
			listBox = this.m_A;
			if (listBox == null)
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
				listBox.PreviewKeyDown += value2;
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
	}

	public wpfPane()
	{
		base.Loaded += wpfPane_Loaded;
		this.m_A = 700.0;
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
		InitializeComponent();
		Pane.TaskPane = this;
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
		this.m_A?.Invoke(this, new PropertyChangedEventArgs(A));
	}

	private void wpfPane_Loaded(object sender, RoutedEventArgs e)
	{
		chkErrors.IsChecked = N.Settings.ProofingShowErrors;
		chkWarnings.IsChecked = N.Settings.ProofingShowWarnings;
		chkMessages.IsChecked = N.Settings.ProofingShowMessages;
		chkErrors.Checked += ToggleErrors;
		chkErrors.Unchecked += ToggleErrors;
		chkWarnings.Checked += ToggleWarnings;
		chkWarnings.Unchecked += ToggleWarnings;
		chkMessages.Checked += ToggleMessages;
		chkMessages.Unchecked += ToggleMessages;
		radErrorsView.Checked += ShowErrors;
		radInfoView.Checked += ShowInfo;
		radSettingsView.Checked += ShowSettings;
	}

	private void btnRefresh_Click(object sender, RoutedEventArgs e)
	{
		A();
	}

	private void A()
	{
		Pane.ActiveItem = null;
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
			base.Dispatcher.Invoke(DispatcherPriority.Background, (ThreadStart)([SpecialName] () =>
			{
			}));
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
			stkLoading.Visibility = Visibility.Collapsed;
			lbxResults.Visibility = Visibility.Visible;
		});
		Pane.CloseCallout();
	}

	public void ShowLoadingState()
	{
		lbxResults.Visibility = Visibility.Collapsed;
		stkLoading.Visibility = Visibility.Visible;
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
			B();
		});
	}

	private void B()
	{
		Microsoft.Office.Interop.Word.Document activeDocument = PC.A.Application.ActiveDocument;
		_ = Pane.WholeDocument;
		B(XC.A(38789));
		Main.Analysis = new Analysis(activeDocument);
		B(XC.A(38832));
		Analysis analysis = Main.Analysis;
		_ = activeDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.StoryType;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = activeDocument.StoryRanges.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				Range range = (Range)enumerator.Current;
				do
				{
					A(range, Main.Analysis);
					try
					{
						WdStoryType storyType = range.StoryType;
						if ((uint)(storyType - 6) <= 5u && range.ShapeRange.Count > 0)
						{
							try
							{
								enumerator2 = range.ShapeRange.GetEnumerator();
								while (enumerator2.MoveNext())
								{
									if (((Microsoft.Office.Interop.Word.Shape)enumerator2.Current).TextFrame2.HasText != MsoTriState.msoTrue)
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
									if (1 == 0)
									{
										/*OpCode not supported: LdMemberToken*/;
									}
									A(range, Main.Analysis);
								}
							}
							finally
							{
								if (enumerator2 is IDisposable)
								{
									while (true)
									{
										switch (5)
										{
										case 0:
											continue;
										}
										(enumerator2 as IDisposable).Dispose();
										break;
									}
								}
							}
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					range = range.NextStoryRange;
				}
				while (range != null);
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
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		if (analysis.Errors.Count > 0)
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
			try
			{
				Analysis analysis2 = analysis;
				IOrderedEnumerable<BaseError> source = analysis.Errors.OrderBy([SpecialName] (BaseError A) => A.PageNumber);
				Func<BaseError, float> keySelector;
				if (_Closure_0024__.A == null)
				{
					keySelector = (_Closure_0024__.A = [SpecialName] (BaseError A) =>
					{
						if (A.Ranges != null)
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
									return Conversions.ToSingle(A.Ranges[0].get_Information(WdInformation.wdVerticalPositionRelativeToPage));
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
						switch (2)
						{
						case 0:
							continue;
						}
						break;
					}
					keySelector = _Closure_0024__.A;
				}
				IOrderedEnumerable<BaseError> source2 = source.ThenBy(keySelector);
				Func<BaseError, float> keySelector2;
				if (_Closure_0024__.B == null)
				{
					float num = default(float);
					keySelector2 = (_Closure_0024__.B = [SpecialName] (BaseError A) => (A.Ranges != null) ? Conversions.ToSingle(A.Ranges[0].get_Information(WdInformation.wdHorizontalPositionRelativeToPage)) : num);
				}
				else
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
					keySelector2 = _Closure_0024__.B;
				}
				IOrderedEnumerable<BaseError> source3 = source2.ThenBy(keySelector2);
				Func<BaseError, float> keySelector3;
				if (_Closure_0024__.C == null)
				{
					float num = default(float);
					keySelector3 = (_Closure_0024__.C = [SpecialName] (BaseError A) => (((BaseError)A).TextRanges != null) ? ((BaseError)A).TextRanges[0].BoundTop : num);
				}
				else
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
					keySelector3 = _Closure_0024__.C;
				}
				IOrderedEnumerable<BaseError> source4 = source3.ThenBy(keySelector3);
				Func<BaseError, float> keySelector4;
				if (_Closure_0024__.D == null)
				{
					keySelector4 = (_Closure_0024__.D = [SpecialName] (BaseError A) =>
					{
						if (((BaseError)A).TextRanges != null)
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
									return ((BaseError)A).TextRanges[0].BoundLeft;
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
						switch (3)
						{
						case 0:
							continue;
						}
						break;
					}
					keySelector4 = _Closure_0024__.D;
				}
				analysis2.Errors = source4.ThenBy(keySelector4).ThenBy([SpecialName] (BaseError A) => (int)((BaseError)A).Severity).ToList();
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			B(XC.A(38873));
			A(analysis.Errors);
		}
		else
		{
			B(XC.A(38910));
			Forms.InfoMessage(XC.A(38939));
		}
		analysis = null;
		activeDocument = null;
	}

	private void A(Range A, Analysis B)
	{
		this.A(A.ShapeRange, B);
		this.A(A.InlineShapes, B);
		this.A(A.Tables, B);
		IEnumerator enumerator = A.Paragraphs.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Paragraph paragraph = (Paragraph)enumerator.Current;
				this.B(B, paragraph.Range);
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
				return;
			}
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	private void A(Microsoft.Office.Interop.Word.ShapeRange A, Analysis B)
	{
		IEnumerator enumerator = A.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Word.Shape a = (Microsoft.Office.Interop.Word.Shape)enumerator.Current;
				this.A(a, B);
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
				return;
			}
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	private void A(InlineShapes A, Analysis B)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				InlineShape a = (InlineShape)enumerator.Current;
				this.A(a, B);
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
				return;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	private void A(Microsoft.Office.Interop.Word.Tables A, Analysis B)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Table a = (Table)enumerator.Current;
				this.A(a, B);
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
				return;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	private void A(Microsoft.Office.Interop.Word.Shape A, Analysis B)
	{
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		//IL_0151: Unknown result type (might be due to invalid IL or missing references)
		//IL_0156: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a4: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ce: Unknown result type (might be due to invalid IL or missing references)
		//IL_008f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0094: Unknown result type (might be due to invalid IL or missing references)
		//IL_0096: Unknown result type (might be due to invalid IL or missing references)
		//IL_00dc: Unknown result type (might be due to invalid IL or missing references)
		//IL_00eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_0109: Unknown result type (might be due to invalid IL or missing references)
		//IL_0118: Unknown result type (might be due to invalid IL or missing references)
		//IL_0127: Unknown result type (might be due to invalid IL or missing references)
		//IL_0134: Unknown result type (might be due to invalid IL or missing references)
		//IL_0143: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cd: Unknown result type (might be due to invalid IL or missing references)
		Microsoft.Office.Interop.Word.Shape shape = A;
		if (A.Type != MsoShapeType.msoGroup)
		{
			if ((int)((Settings)B.Options).ColorPalette != 0)
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
				MsoShapeType type = A.Type;
				if (type <= MsoShapeType.msoChart)
				{
					if (type != MsoShapeType.msoAutoShape && type != MsoShapeType.msoChart)
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
						goto IL_014b;
					}
				}
				else if (type != MsoShapeType.msoTextBox)
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
					if (type != MsoShapeType.msoTable)
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
						if (type != MsoShapeType.msoSmartArt)
						{
							goto IL_014b;
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
				Severity colorPalette = ((Settings)B.Options).ColorPalette;
				Checks checks = B.Checks;
				if (A.HasChart == MsoTriState.msoFalse)
				{
					ShapeColors.FillColor(A, ((Checks)checks).PaletteColors, colorPalette);
					ShapeColors.BorderColor(A, ((Checks)checks).PaletteColors, colorPalette);
					ShapeColors.FontColor(A, ((Checks)checks).PaletteColors, colorPalette);
				}
				else
				{
					ChartColors.ChartArea(A, ((Checks)checks).PaletteColors, colorPalette);
					ChartColors.PlotArea(A, ((Checks)checks).PaletteColors, colorPalette);
					ChartColors.Series(A, ((Checks)checks).PaletteColors, colorPalette);
					ChartColors.Axes(A, ((Checks)checks).PaletteColors, colorPalette);
					ChartColors.Title(A, ((Checks)checks).PaletteColors, colorPalette);
					ChartColors.Legend(A, ((Checks)checks).PaletteColors, colorPalette);
					ChartColors.DataTable(A, ((Checks)checks).PaletteColors, colorPalette);
					ChartColors.Groups(A, ((Checks)checks).PaletteColors, colorPalette);
				}
				checks = null;
			}
			goto IL_014b;
		}
		int count = shape.GroupItems.Count;
		for (int i = 1; i <= count; i = checked(i + 1))
		{
			GroupShapes groupItems = shape.GroupItems;
			object Index = i;
			this.A(groupItems[ref Index], B);
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
		goto IL_035a;
		IL_0305:
		Images.Distortion(A);
		goto IL_035a;
		IL_014b:
		if ((int)((Settings)B.Options).FillTransparency != 0)
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
			ShapeColors.FillTransparency(A);
		}
		Checks checks2 = B.Checks;
		if (checks2.ShapeOutsideMargins != null)
		{
			checks2.ShapeOutsideMargins.Check(A.Anchor.Document, A);
		}
		checks2 = null;
		if ((int)((Settings)B.Options).CheckLinks != 0)
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
			Miscellaneous.CheckMacabacusLink(A);
		}
		if (A.HasChart == MsoTriState.msoTrue)
		{
			if ((int)((Settings)B.Options).CheckChartElements != 0)
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
				ChartElements.Legend(A);
				ChartElements.MissingDataLabels(A);
				ChartElements.DataLabelConsistency(A);
				ChartElements.DataLabelFormat(A);
			}
		}
		else if (A.HasSmartArt == MsoTriState.msoTrue)
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
			try
			{
				IEnumerator enumerator = A.SmartArt.Nodes.GetEnumerator();
				try
				{
					while (enumerator.MoveNext())
					{
						SmartArtNode smartArtNode = (SmartArtNode)enumerator.Current;
						if (smartArtNode.TextFrame2.HasText != MsoTriState.msoTrue)
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
						this.A(B, A, smartArtNode.TextFrame2.TextRange);
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_027d;
						}
						continue;
						end_IL_027d:
						break;
					}
				}
				finally
				{
					IDisposable disposable = enumerator as IDisposable;
					if (disposable != null)
					{
						disposable.Dispose();
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
		else
		{
			MsoShapeType type2 = A.Type;
			if (type2 <= MsoShapeType.msoLinkedPicture)
			{
				if (type2 != MsoShapeType.msoAutoShape)
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
					if (type2 == MsoShapeType.msoLinkedPicture)
					{
						goto IL_0305;
					}
				}
			}
			else
			{
				if (type2 == MsoShapeType.msoPicture)
				{
					goto IL_0305;
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
				if (type2 != MsoShapeType.msoTextBox)
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
			}
		}
		goto IL_035a;
		IL_035a:
		shape = null;
	}

	private void A(InlineShape A, Analysis B)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0147: Unknown result type (might be due to invalid IL or missing references)
		//IL_014c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0196: Unknown result type (might be due to invalid IL or missing references)
		//IL_019b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0077: Unknown result type (might be due to invalid IL or missing references)
		//IL_007c: Unknown result type (might be due to invalid IL or missing references)
		//IL_007e: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00df: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ee: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_010a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0119: Unknown result type (might be due to invalid IL or missing references)
		//IL_0128: Unknown result type (might be due to invalid IL or missing references)
		//IL_0137: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c3: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ca: Unknown result type (might be due to invalid IL or missing references)
		//IL_01cf: Unknown result type (might be due to invalid IL or missing references)
		if ((int)((Settings)B.Options).ColorPalette != 0)
		{
			WdInlineShapeType type = A.Type;
			if (type <= WdInlineShapeType.wdInlineShapePicture)
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
				if (type != WdInlineShapeType.wdInlineShapeEmbeddedOLEObject && type != WdInlineShapeType.wdInlineShapePicture)
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
					goto IL_013f;
				}
			}
			else if (type != (WdInlineShapeType)17)
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
				if (type != (WdInlineShapeType)19)
				{
					if (type != (WdInlineShapeType)24)
					{
						goto IL_013f;
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
			Severity colorPalette = ((Settings)B.Options).ColorPalette;
			Checks checks = B.Checks;
			if (A.HasChart == MsoTriState.msoFalse)
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
				ShapeColors.FillColor(A, ((Checks)checks).PaletteColors, colorPalette);
				ShapeColors.BorderColor(A, ((Checks)checks).PaletteColors, colorPalette);
				ShapeColors.FontColor(A, ((Checks)checks).PaletteColors, colorPalette);
			}
			else
			{
				ChartColors.ChartArea(A, ((Checks)checks).PaletteColors, colorPalette);
				ChartColors.PlotArea(A, ((Checks)checks).PaletteColors, colorPalette);
				ChartColors.Series(A, ((Checks)checks).PaletteColors, colorPalette);
				ChartColors.Axes(A, ((Checks)checks).PaletteColors, colorPalette);
				ChartColors.Title(A, ((Checks)checks).PaletteColors, colorPalette);
				ChartColors.Legend(A, ((Checks)checks).PaletteColors, colorPalette);
				ChartColors.DataTable(A, ((Checks)checks).PaletteColors, colorPalette);
				ChartColors.Groups(A, ((Checks)checks).PaletteColors, colorPalette);
			}
			checks = null;
		}
		goto IL_013f;
		IL_013f:
		if ((int)((Settings)B.Options).FillTransparency != 0)
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
			ShapeColors.FillTransparency(A);
		}
		Checks checks2 = B.Checks;
		if (checks2.ShapeOutsideMargins != null)
		{
			checks2.ShapeOutsideMargins.Check(A.Range.Document, A);
		}
		checks2 = null;
		if ((int)((Settings)B.Options).CheckLinks != 0)
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
			Miscellaneous.CheckMacabacusLink(A);
		}
		if (A.HasChart == MsoTriState.msoTrue)
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
			if ((int)((Settings)B.Options).CheckChartElements != 0)
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
				ChartElements.Legend((Microsoft.Office.Interop.Word.Shape)A);
				ChartElements.MissingDataLabels((Microsoft.Office.Interop.Word.Shape)A);
				ChartElements.DataLabelConsistency((Microsoft.Office.Interop.Word.Shape)A);
				ChartElements.DataLabelFormat((Microsoft.Office.Interop.Word.Shape)A);
			}
		}
		else if (A.HasSmartArt == MsoTriState.msoTrue)
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
				foreach (SmartArtNode node in A.SmartArt.Nodes)
				{
					if (node.TextFrame2.HasText == MsoTriState.msoTrue)
					{
						this.A(B, A, node.TextFrame2.TextRange);
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
		else
		{
			WdInlineShapeType type2 = A.Type;
			if (type2 <= WdInlineShapeType.wdInlineShapeOWSAnchor)
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
				if (type2 != WdInlineShapeType.wdInlineShapeEmbeddedOLEObject)
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
					if (type2 == WdInlineShapeType.wdInlineShapeOWSAnchor)
					{
						goto IL_0312;
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
				if (type2 == WdInlineShapeType.wdInlineShapeDiagram)
				{
					goto IL_0312;
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
				if (type2 != (WdInlineShapeType)17)
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
			}
		}
		goto IL_031d;
		IL_031d:
		_ = null;
		return;
		IL_0312:
		Images.Distortion((Microsoft.Office.Interop.Word.Shape)A);
		goto IL_031d;
	}

	private void A(Table A, Analysis B)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		//IL_004a: Unknown result type (might be due to invalid IL or missing references)
		List<CellPadding> list = new List<CellPadding>();
		List<Range> list2 = new List<Range>();
		if ((int)((Settings)B.Options).ColorPalette != 0)
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
			Analysis analysis = B;
			Macabacus_Word.Proofing.Check.Tables.Colors(A, ((Checks)analysis.Checks).PaletteColors, ((Settings)analysis.Options).ColorPalette);
			analysis = null;
		}
		Table table = A;
		int count = table.Rows.Count;
		int count2 = table.Columns.Count;
		int num = count;
		checked
		{
			for (int i = 1; i <= num; i++)
			{
				int num2 = count2;
				for (int j = 1; j <= num2; j++)
				{
					try
					{
						Cell cell = table.Cell(i, j);
						try
						{
							list.Add(new CellPadding
							{
								Top = cell.TopPadding,
								Right = cell.RightPadding,
								Bottom = cell.BottomPadding,
								Left = cell.LeftPadding
							});
							list2.Add(A.Cell(i, j).Range);
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						if (cell.Range.Text.Length > 0)
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
							this.B(B, cell.Range);
						}
						cell = null;
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
				}
			}
			table = null;
			Macabacus_Word.Proofing.Check.Tables.CellMargins(A, list, list2);
			list = null;
			list2 = null;
		}
	}

	private void B(Analysis A, Range B)
	{
		Checks checks = A.Checks;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = B.Paragraphs.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Paragraph paragraph = (Paragraph)enumerator.Current;
				string strText = Text.PrintableText(paragraph.Range.Text);
				foreach (BaseTextCheck textCheck in checks.TextChecks)
				{
					textCheck.Check(paragraph.Range, strText);
				}
				if (checks.MultipleFontFamilies != null)
				{
					checks.MultipleFontFamilies.Check(paragraph.Range);
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				checks.ProofingLanguage.Check(paragraph.Range);
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_00d4;
				}
				continue;
				end_IL_00d4:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		checks = null;
	}

	private void A(Analysis A, object B, TextRange2 C)
	{
		Checks checks = A.Checks;
		int count = C.get_Paragraphs(-1, -1).Count;
		for (int i = 1; i <= count; i = checked(i + 1))
		{
			TextRange2 textRange = C.get_Paragraphs(i, -1);
			string strText = Text.PrintableText(textRange.Text);
			using (List<BaseTextCheck>.Enumerator enumerator = checks.TextChecks.GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					enumerator.Current.Check(RuntimeHelpers.GetObjectValue(B), textRange, strText);
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
					break;
				}
			}
			if (checks.MultipleFontFamilies != null)
			{
				checks.MultipleFontFamilies.Check(RuntimeHelpers.GetObjectValue(B), textRange);
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
			checks.ProofingLanguage.Check((Range)textRange);
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			checks = null;
			return;
		}
	}

	private void A(List<BaseError> A)
	{
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		List<BaseError> list = A.Where([SpecialName] (BaseError baseError) => baseError.Range == null).ToList();
		Func<BaseError, bool> predicate;
		if (_Closure_0024__.B == null)
		{
			predicate = (_Closure_0024__.B = [SpecialName] (BaseError baseError) => baseError.Range != null);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			predicate = _Closure_0024__.B;
		}
		List<BaseError> list2 = A.Where(predicate).ToList();
		List<BaseError> source = list2;
		Func<BaseError, int> keySelector;
		if (_Closure_0024__.C == null)
		{
			keySelector = (_Closure_0024__.C = [SpecialName] (BaseError baseError) => baseError.PageNumber);
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
			keySelector = _Closure_0024__.C;
		}
		Func<BaseError, BaseError> elementSelector;
		if (_Closure_0024__.A == null)
		{
			elementSelector = (_Closure_0024__.A = [SpecialName] (BaseError result) => result);
		}
		else
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
			elementSelector = _Closure_0024__.A;
		}
		Func<int, IEnumerable<BaseError>, F<int, IEnumerable<BaseError>>> resultSelector;
		if (_Closure_0024__.A == null)
		{
			resultSelector = (_Closure_0024__.A = [SpecialName] (int a, IEnumerable<BaseError> B) => new F<int, IEnumerable<BaseError>>(a, B));
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
			resultSelector = _Closure_0024__.A;
		}
		IEnumerable<F<int, IEnumerable<BaseError>>> source2 = source.GroupBy(keySelector, elementSelector, resultSelector);
		Func<F<int, IEnumerable<BaseError>>, int> keySelector2;
		if (_Closure_0024__.A == null)
		{
			keySelector2 = (_Closure_0024__.A = [SpecialName] (F<int, IEnumerable<BaseError>> f) => f.idx);
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
			keySelector2 = _Closure_0024__.A;
		}
		IOrderedEnumerable<F<int, IEnumerable<BaseError>>> source3 = source2.OrderBy(keySelector2);
		Func<F<int, IEnumerable<BaseError>>, G<int, int>> selector;
		if (_Closure_0024__.A == null)
		{
			selector = (_Closure_0024__.A = [SpecialName] (F<int, IEnumerable<BaseError>> f) => new G<int, int>(f.idx, f.Group.Count()));
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
			selector = _Closure_0024__.A;
		}
		IEnumerable<G<int, int>> enumerable = source3.Select(selector);
		foreach (G<int, int> item in enumerable)
		{
			dictionary.Add(item.PageNumber, item.Count);
		}
		AllItems = new ObservableCollection<BaseError>();
		if (list.Count > 0)
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
			string text = Guid.NewGuid().ToString();
			int num = this.A(list);
			int num2 = B(list);
			int num3 = C(list);
			foreach (BaseError item2 in list)
			{
				((BaseError)item2).SetVisualProperties(XC.A(39020), text, num, num2, num3);
				AllItems.Add(item2);
			}
		}
		using (Dictionary<int, int>.Enumerator enumerator3 = dictionary.GetEnumerator())
		{
			JC jC = default(JC);
			while (enumerator3.MoveNext())
			{
				KeyValuePair<int, int> current3 = enumerator3.Current;
				jC = new JC(jC);
				jC.A = current3.Key;
				list2 = A.Where(jC.A).ToList();
				string text = Guid.NewGuid().ToString();
				int num = this.A(list2);
				int num2 = B(list2);
				int num3 = C(list2);
				string text2 = XC.A(39037) + current3.Key;
				using List<BaseError>.Enumerator enumerator4 = list2.GetEnumerator();
				while (enumerator4.MoveNext())
				{
					BaseError current4 = enumerator4.Current;
					((BaseError)current4).SetVisualProperties(text2, text, num, num2, num3);
					AllItems.Add(current4);
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_0365;
					}
					continue;
					end_IL_0365:
					break;
				}
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_038d;
				}
				continue;
				end_IL_038d:
				break;
			}
		}
		lbxResults.SelectionChanged -= ListBoxSelectionChanged;
		SourceCollection = CollectionViewSource.GetDefaultView(AllItems);
		SourceCollection.GroupDescriptions.Add(new PropertyGroupDescription(XC.A(8561)));
		SourceCollection.Filter = this.A;
		lbxResults.SelectionChanged += ListBoxSelectionChanged;
		list = null;
		list2 = null;
		dictionary = null;
	}

	private int A(List<BaseError> A)
	{
		Func<BaseError, bool> predicate;
		if (_Closure_0024__.C == null)
		{
			predicate = (_Closure_0024__.C = [SpecialName] (BaseError baseError) => (int)((BaseError)baseError).Severity == 3);
		}
		else
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
			predicate = _Closure_0024__.C;
		}
		return A.Where(predicate).Count();
	}

	private int B(List<BaseError> A)
	{
		return A.Where([SpecialName] (BaseError baseError) => (int)((BaseError)baseError).Severity == 2).Count();
	}

	private int C(List<BaseError> A)
	{
		Func<BaseError, bool> predicate;
		if (_Closure_0024__.E == null)
		{
			predicate = (_Closure_0024__.E = [SpecialName] (BaseError baseError) => (int)((BaseError)baseError).Severity == 1);
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
			predicate = _Closure_0024__.E;
		}
		return A.Where(predicate).Count();
	}

	private void lbxResults_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
	{
		Key key = e.Key;
		if (key <= Key.Space)
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
					switch (key)
					{
					default:
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								return;
							}
						}
					case Key.Escape:
						if (Callout.Dialog != null)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									Callout.Dialog.Close();
									e.Handled = true;
									return;
								}
							}
						}
						break;
					case Key.Space:
						if (lbxResults.SelectedIndex > -1)
						{
							C((BaseError)lbxResults.SelectedItem);
						}
						break;
					}
					return;
				}
			}
		}
		if ((uint)(key - 23) > 3u)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (key != Key.Delete)
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
					if (lbxResults.SelectedIndex > -1)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								RemoveItemAndNavigate((BaseError)lbxResults.SelectedItem);
								return;
							}
						}
					}
					return;
				}
			}
		}
		if (e.IsRepeat)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					if (Callout.Dialog != null)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								Callout.Dialog.Top = -10000.0;
								Callout.RemoveMarchingAnts();
								return;
							}
						}
					}
					return;
				}
			}
		}
		lbxResults.KeyUp += NavKeyUp;
	}

	private void NavKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
	{
		lbxResults.KeyUp -= NavKeyUp;
		SelectionChanged();
		e.Handled = true;
	}

	private void ListBoxItemLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
	{
		try
		{
			if (e.NewFocus is ListBoxItem)
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
				if (e.NewFocus is ToggleButton)
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
					if (Callout.DoNotClose)
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
						if (Callout.Dialog.IsMouseOver)
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
							if (CollapseAnimationRunning)
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
								if (Callout.Dialog != null)
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
									Callout.Dialog.Close();
								}
								System.Windows.Controls.ListBox listBox = lbxResults;
								if (listBox.SelectedIndex > -1)
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
									if (((BaseError)(BaseError)listBox.SelectedItem).IsFixed)
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
										RemoveItem((BaseError)listBox.SelectedItem, blnAnimate: true);
									}
								}
								listBox.SelectedIndex = -1;
								listBox = null;
								return;
							}
						}
					}
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

	private void ListBoxSelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		SelectionChanged();
	}

	public void SelectionChanged()
	{
		if (lbxResults.SelectedIndex > -1)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					BaseError baseError = (BaseError)lbxResults.SelectedItem;
					if (Pane.ActiveItem != null && ((BaseError)Pane.ActiveItem).IsFixed)
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
						RemoveItem(Pane.ActiveItem, blnAnimate: true);
					}
					Pane.ActiveItem = baseError;
					if (!System.Windows.Input.Keyboard.IsKeyDown(Key.Down))
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
						if (!System.Windows.Input.Keyboard.IsKeyDown(Key.Up) && !System.Windows.Input.Keyboard.IsKeyDown(Key.Right))
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
							if (!System.Windows.Input.Keyboard.IsKeyDown(Key.Left) && !System.Windows.Input.Keyboard.IsKeyDown(Key.Next))
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
								if (!System.Windows.Input.Keyboard.IsKeyDown(Key.Prior))
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
									try
									{
										try
										{
											this.m_A.Stop();
										}
										catch (Exception ex)
										{
											ProjectData.SetProjectError(ex);
											Exception ex2 = ex;
											ProjectData.ClearProjectError();
										}
										this.m_A = new DispatcherTimer();
										this.m_A.Interval = TimeSpan.FromMilliseconds(this.m_A);
										this.m_A.Tick += [SpecialName] (object obj, EventArgs ev) =>
										{
											this.m_A.Stop();
											((BaseError)Pane.ActiveItem).FontWeight = FontWeights.Normal;
										};
										this.m_A.Start();
									}
									catch (Exception ex3)
									{
										ProjectData.SetProjectError(ex3);
										Exception ex4 = ex3;
										ProjectData.ClearProjectError();
									}
									A(baseError);
								}
							}
						}
					}
					baseError = null;
					return;
				}
				}
			}
		}
		Pane.ActiveItem = null;
	}

	private void A(BaseError A)
	{
		if (Callout.Dialog != null)
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
					Callout.Dialog.GoToError(A);
					return;
				}
			}
		}
		wpfCallout obj = new wpfCallout();
		obj.Top = -10000.0;
		obj.ShowActivated = false;
		obj.Show();
		obj.GoToError(A);
		_ = null;
	}

	private void B(BaseError A)
	{
		if (lbxResults.SelectedItem == A)
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
			lbxResults.SelectedItem = A;
			return;
		}
	}

	private void ToggleErrors(object sender, RoutedEventArgs e)
	{
		C();
		N.Settings.ProofingShowErrors = chkErrors.IsChecked.Value;
	}

	private void ToggleWarnings(object sender, RoutedEventArgs e)
	{
		C();
		N.Settings.ProofingShowWarnings = chkWarnings.IsChecked.Value;
	}

	private void ToggleMessages(object sender, RoutedEventArgs e)
	{
		C();
		N.Settings.ProofingShowMessages = chkMessages.IsChecked.Value;
	}

	private void C()
	{
		D();
		if (lbxResults.SelectedIndex <= -1)
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
			Pane.FocusListBoxItem(lbxResults);
			return;
		}
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
			D();
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
			switch (6)
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

	private void SearchTextBoxGotFocus(object sender, RoutedEventArgs e)
	{
		if (txtSearch.Text.Length > 0)
		{
			txtSearch.SelectAll();
		}
	}

	private void SearchTextBoxLostFocus(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Return)
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
			if (txtSearch.Text.Length <= 0)
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
				txtSearch.SelectAll();
				D();
				return;
			}
		}
	}

	private void D()
	{
		lbxResults.SelectionChanged -= ListBoxSelectionChanged;
		Pane.CloseCallout();
		Pane.ActiveItem = null;
		SourceCollection = CollectionViewSource.GetDefaultView(lbxResults.ItemsSource);
		SourceCollection.Filter = A;
		SourceCollection.Refresh();
		lbxResults.SelectionChanged += ListBoxSelectionChanged;
	}

	private bool A(object A)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0010: Invalid comparison between Unknown and I4
		//IL_0050: Unknown result type (might be due to invalid IL or missing references)
		//IL_0055: Unknown result type (might be due to invalid IL or missing references)
		//IL_0058: Invalid comparison between Unknown and I4
		//IL_0096: Unknown result type (might be due to invalid IL or missing references)
		//IL_009b: Unknown result type (might be due to invalid IL or missing references)
		//IL_009e: Invalid comparison between Unknown and I4
		BaseError baseError = (BaseError)A;
		if ((int)((BaseError)baseError).Severity == 3 && chkErrors.IsChecked == true)
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
					return this.A(baseError);
				}
			}
		}
		if ((int)((BaseError)baseError).Severity == 2)
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
			if (chkWarnings.IsChecked == true)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						return this.A(baseError);
					}
				}
			}
		}
		if ((int)((BaseError)baseError).Severity == 1)
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
			if (chkMessages.IsChecked == true)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						return this.A(baseError);
					}
				}
			}
		}
		return false;
	}

	private bool A(BaseError A)
	{
		string text = txtSearch.Text.ToLower();
		if (text.Length == 0)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return true;
				}
			}
		}
		if (!((BaseError)A).Title.ToLower().Contains(text))
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
			if (!((BaseError)A).Subtitle.ToLower().Contains(text))
			{
				return false;
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
		return true;
	}

	private void FixButtonClicked(object sender, RoutedEventArgs e)
	{
		C((BaseError)((System.Windows.Controls.Button)sender).DataContext);
	}

	private void C(BaseError A)
	{
		B(A);
		Fixes.DefaultFixButtonClicked(A);
	}

	private void ShowFixOptions(object sender, RoutedEventArgs e)
	{
		BaseError baseError = (BaseError)((ToggleButton)sender).DataContext;
		B(baseError);
		Fixes.ShowOptions(baseError, (ToggleButton)sender);
		baseError = null;
	}

	private void A(ListBoxItem A)
	{
		CollapseAnimationRunning = true;
		DoubleAnimation collapseAnimation = Pane.GetCollapseAnimation(A);
		collapseAnimation.Completed += CollapseComplete;
		Pane.CollapseListBoxItem(A, collapseAnimation);
		collapseAnimation = null;
	}

	private void CollapseComplete(object sender, EventArgs e)
	{
		E();
		CollapseAnimationRunning = false;
	}

	private void E()
	{
		F();
		AllItems.Remove(ItemsQueuedToRemove[0]);
		Main.Analysis.Errors.Remove(ItemsQueuedToRemove[0]);
		ItemsQueuedToRemove.RemoveAt(0);
		if (NavigateAfterRemoveIndex > -1)
		{
			lbxResults.SelectedIndex = NavigateAfterRemoveIndex;
			Pane.FocusListBoxItem(lbxResults);
			NavigateAfterRemoveIndex = -1;
		}
	}

	private void F()
	{
		//IL_0053: Unknown result type (might be due to invalid IL or missing references)
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005b: Invalid comparison between Unknown and I4
		//IL_0078: Unknown result type (might be due to invalid IL or missing references)
		//IL_007d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0080: Invalid comparison between Unknown and I4
		List<BaseError> list = AllItems.Where([SpecialName] (BaseError A) => Operators.CompareString(((BaseError)A).Guid, ((BaseError)ItemsQueuedToRemove[0]).Guid, TextCompare: false) == 0).ToList();
		BaseError baseError = ItemsQueuedToRemove[0];
		int num = ((BaseError)baseError).ErrorsCount;
		int num2 = ((BaseError)baseError).WarningsCount;
		int num3 = ((BaseError)baseError).MessagesCount;
		if ((int)((BaseError)baseError).Severity == 2)
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
			num2 = checked(num2 - 1);
		}
		else if ((int)((BaseError)baseError).Severity == 3)
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
			num = checked(num - 1);
		}
		else
		{
			num3 = checked(num3 - 1);
		}
		baseError = null;
		using (List<BaseError>.Enumerator enumerator = list.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				BaseError current = enumerator.Current;
				((BaseError)current).ErrorsCount = num;
				((BaseError)current).WarningsCount = num2;
				((BaseError)current).MessagesCount = num3;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_00cb;
				}
				continue;
				end_IL_00cb:
				break;
			}
		}
		list = null;
	}

	private void RemoveButtonCicked(object sender, RoutedEventArgs e)
	{
		RemoveItemAndNavigate((BaseError)((System.Windows.Controls.Button)sender).DataContext);
	}

	public void RemoveItemAndNavigate(BaseError err)
	{
		NavigateAfterRemoveIndex = lbxResults.SelectedIndex;
		RemoveItem(err, blnAnimate: true);
	}

	public void RemoveItem(BaseError itm, bool blnAnimate)
	{
		checked
		{
			if (((BaseError)itm).IsFixed)
			{
				BaseError baseError = itm;
				try
				{
					if (baseError is Macabacus_Word.Proofing.Errors.AbbreviationSpacing)
					{
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
							using (List<BaseError>.Enumerator enumerator = Main.Analysis.Errors.GetEnumerator())
							{
								while (true)
								{
									if (enumerator.MoveNext())
									{
										BaseError current = enumerator.Current;
										if (!(current is Macabacus_Word.Proofing.Errors.AbbreviationMillions))
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
											if (!(current is AbbreviationBillions))
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
										}
										if (((BaseError)current).TextRanges[0] == null)
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
										if (current.Shape != baseError.Shape)
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
										TextRange2 textRange = ((BaseError)baseError).TextRanges[0];
										if (((BaseError)current).TextRanges[0].Start >= textRange.Start)
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
											if (((BaseError)current).TextRanges[0].Start < textRange.Start + textRange.Length)
											{
												while (true)
												{
													switch (3)
													{
													case 0:
														continue;
													}
													List<string> list = new List<string>(new string[12]
													{
														XC.A(39048),
														XC.A(39053),
														XC.A(39058),
														XC.A(39063),
														XC.A(39068),
														XC.A(39071),
														XC.A(39074),
														XC.A(39081),
														XC.A(39086),
														XC.A(39091),
														XC.A(39094),
														XC.A(39097)
													});
													Match match = new Regex(XC.A(25505) + Strings.Join(list.ToArray(), XC.A(19662)) + XC.A(20696), RegexOptions.None).Match(textRange.Text);
													if (match != null)
													{
														Match match2 = match;
														((BaseError)current).TextRanges[0] = textRange.get_Characters(match2.Index + 1, match2.Length);
														match2 = null;
														match = null;
													}
													list = null;
													break;
												}
												break;
											}
										}
										textRange = null;
										continue;
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											break;
										default:
											goto end_IL_026f;
										}
										continue;
										end_IL_026f:
										break;
									}
									break;
								}
							}
							break;
						}
					}
					else
					{
						if (baseError is Macabacus_Word.Proofing.Errors.AbbreviationMillions)
						{
							goto IL_02aa;
						}
						if (baseError is AbbreviationBillions)
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
							goto IL_02aa;
						}
					}
					goto end_IL_000d;
					IL_02aa:
					foreach (BaseError error in Main.Analysis.Errors)
					{
						if (!(error is Macabacus_Word.Proofing.Errors.AbbreviationSpacing) || ((BaseError)error).TextRanges[0] == null || error.Shape != baseError.Shape)
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
						TextRange2 textRange2 = ((BaseError)baseError).TextRanges[0];
						if (textRange2.Start >= ((BaseError)error).TextRanges[0].Start)
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
							if (textRange2.Start < ((BaseError)error).TextRanges[0].Start + ((BaseError)error).TextRanges[0].Length)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									List<string> list2 = new List<string>(new string[12]
									{
										XC.A(39048),
										XC.A(39053),
										XC.A(39058),
										XC.A(39063),
										XC.A(39068),
										XC.A(39071),
										XC.A(39074),
										XC.A(39081),
										XC.A(39086),
										XC.A(39091),
										XC.A(39094),
										XC.A(39097)
									});
									((BaseError)error).ReplacementText[0] = Regex.Replace(((BaseError)error).ReplacementText[0], XC.A(39104) + Strings.Join(list2.ToArray(), XC.A(19662)) + XC.A(20696), XC.A(24538) + ((BaseError)baseError).ReplacementText[0]);
									list2 = null;
									break;
								}
								break;
							}
						}
						textRange2 = null;
					}
					end_IL_000d:;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				baseError = null;
			}
			if (ItemsQueuedToRemove == null)
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
				ItemsQueuedToRemove = new List<BaseError>();
			}
			ItemsQueuedToRemove.Add(itm);
			if (blnAnimate)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						A((ListBoxItem)lbxResults.ItemContainerGenerator.ContainerFromItem(itm));
						return;
					}
				}
			}
			E();
		}
	}

	private void ShowErrors(object sender, RoutedEventArgs e)
	{
		conReformat.Visibility = Visibility.Collapsed;
		conSettings.Visibility = Visibility.Collapsed;
		lbxResults.Visibility = Visibility.Visible;
		if (!SettingsDirty)
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
			SettingsView.Save();
			A();
			SettingsDirty = false;
			return;
		}
	}

	private void ShowInfo(object sender, RoutedEventArgs e)
	{
		if (ReformatView == null)
		{
			ReformatView = new wpfReformat(Main.Analysis.Conventions);
		}
		lbxResults.Visibility = Visibility.Collapsed;
		conSettings.Visibility = Visibility.Collapsed;
		conReformat.Visibility = Visibility.Visible;
	}

	private void ShowSettings(object sender, RoutedEventArgs e)
	{
		if (SettingsView == null)
		{
			SettingsView = new wpfSettings();
		}
		lbxResults.Visibility = Visibility.Collapsed;
		conReformat.Visibility = Visibility.Collapsed;
		conSettings.Visibility = Visibility.Visible;
		SettingsDirty = true;
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (this.m_C)
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
			this.m_C = true;
			Uri resourceLocator = new Uri(XC.A(39125), UriKind.Relative);
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
			grdHeader = (Grid)target;
			return;
		}
		if (connectionId == 2)
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
					btnRefresh = (System.Windows.Controls.Button)target;
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
					chkErrors = (System.Windows.Controls.CheckBox)target;
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
					chkWarnings = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkMessages = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			radErrorsView = (System.Windows.Controls.RadioButton)target;
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
					radInfoView = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					radSettingsView = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					txtSearch = (System.Windows.Controls.TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					stkLoading = (StackPanel)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					cvsLoading = (Canvas)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					noFreeze = (RotateTransform)target;
					return;
				}
			}
		}
		if (connectionId == 13)
		{
			txtLoading = (TextBlock)target;
			return;
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
					conReformat = (System.Windows.Controls.ContentControl)target;
					return;
				}
			}
		}
		if (connectionId == 15)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					conSettings = (System.Windows.Controls.ContentControl)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					lbxResults = (System.Windows.Controls.ListBox)target;
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

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 17)
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
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = UIElement.LostKeyboardFocusEvent;
			eventSetter.Handler = new KeyboardFocusChangedEventHandler(ListBoxItemLostKeyboardFocus);
			((Style)target).Setters.Add(eventSetter);
		}
		if (connectionId == 18)
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
			((System.Windows.Controls.Button)target).Click += FixButtonClicked;
		}
		if (connectionId == 19)
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
			((ToggleButton)target).Checked += ShowFixOptions;
		}
		if (connectionId != 20)
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
			((System.Windows.Controls.Button)target).Click += RemoveButtonCicked;
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
	private void G()
	{
		stkLoading.Visibility = Visibility.Collapsed;
		lbxResults.Visibility = Visibility.Visible;
	}

	[SpecialName]
	[CompilerGenerated]
	private void H()
	{
		B();
	}

	[SpecialName]
	[CompilerGenerated]
	private bool B(BaseError A)
	{
		return Operators.CompareString(((BaseError)A).Guid, ((BaseError)ItemsQueuedToRemove[0]).Guid, TextCompare: false) == 0;
	}
}
