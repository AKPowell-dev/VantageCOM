using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
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
using ExcelAddIn1.Audit.Check.Helpers;
using ExcelAddIn1.Audit.Check.Observations;
using ExcelAddIn1.Sheets;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.Proofing.UI;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.UI;

[DesignerGenerated]
public sealed class wpfAudit : System.Windows.Controls.UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Observation, string> A;

		public static Func<Observation, double> A;

		public static Func<Observation, string> B;

		public static Func<Observation, int> A;

		public static Func<Observation, object> A;

		public static Func<Observation, object> B;

		public static Func<Observation, string> C;

		public static Func<Observation, bool> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal string A(Observation A)
		{
			return A.Title;
		}

		[SpecialName]
		internal double A(Observation A)
		{
			return (double)A.Severity;
		}

		[SpecialName]
		internal string B(Observation A)
		{
			return A.Title;
		}

		[SpecialName]
		internal int A(Observation A)
		{
			return A.SheetIndex;
		}

		[SpecialName]
		internal object A(Observation A)
		{
			if (A.Range == null)
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
						return 0;
					}
				}
			}
			return NewLateBinding.LateGet(A.Range.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41347), new object[0], null, null, null);
		}

		[SpecialName]
		internal object B(Observation A)
		{
			if (A.Range == null)
			{
				return 0;
			}
			return NewLateBinding.LateGet(A.Range.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41354), new object[0], null, null, null);
		}

		[SpecialName]
		internal string C(Observation A)
		{
			return A.Title;
		}

		[SpecialName]
		internal bool A(Observation A)
		{
			return A is FC;
		}
	}

	[CompilerGenerated]
	internal sealed class JC
	{
		public DoWorkEventArgs A;

		public wpfAudit A;

		[SpecialName]
		internal void A()
		{
			this.A.A(ref this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class KC
	{
		public RunWorkerCompletedEventArgs A;

		public wpfAudit A;

		[SpecialName]
		internal void A()
		{
			this.A.IsRunningAnalysis = false;
			if (!this.A.btnCancel.IsEnabled)
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
				this.A.btnCancel.IsEnabled = true;
			}
			this.A.A(this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class LC
	{
		public Dictionary<string, double> A;

		public LC(LC A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal double A(Observation A)
		{
			return this.A[A.Title];
		}
	}

	[CompilerGenerated]
	internal sealed class MC
	{
		public string A;

		[SpecialName]
		internal bool A(string A)
		{
			object obj = A;
			if (obj == null)
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
				obj = "";
			}
			return ((string)obj).ToLower().Contains(this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class NC
	{
		public Observation A;

		public Func<Observation, bool> A;

		public NC(NC A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(Observation A)
		{
			return object.Equals(A.Title, this.A.Title);
		}
	}

	[CompilerGenerated]
	internal sealed class OC
	{
		public object A;

		public OC(OC A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(Observation A)
		{
			return Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(A.Sheet, null, VH.A(48135), new object[0], null, null, null), NewLateBinding.LateGet(this.A, null, VH.A(48135), new object[0], null, null, null), TextCompare: false);
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private readonly double m_A;

	private DispatcherTimer m_A;

	[CompilerGenerated]
	private ctpSuperFind2 m_A;

	[CompilerGenerated]
	private BackgroundWorker m_A;

	private StatusKeeper m_A;

	private ICollectionView m_A;

	[CompilerGenerated]
	private ObservableCollection<Observation> m_A;

	[CompilerGenerated]
	private Observation m_A;

	[CompilerGenerated]
	private Settings m_A;

	[CompilerGenerated]
	private int m_A;

	[CompilerGenerated]
	private bool m_A;

	[CompilerGenerated]
	private List<Observation> m_A;

	[CompilerGenerated]
	private List<Type> m_A;

	[CompilerGenerated]
	private bool m_B;

	[CompilerGenerated]
	private static wpfAudit m_A;

	private bool m_C;

	private bool m_D;

	private bool m_E;

	private System.Windows.Visibility m_A;

	private string m_A;

	private bool m_F;

	private bool m_G;

	private string m_B;

	[CompilerGenerated]
	private string m_C;

	private bool m_H;

	private bool m_I;

	private bool m_J;

	private bool m_K;

	[CompilerGenerated]
	private Range m_A;

	[CompilerGenerated]
	private static (bool YesNo, List<int> SheetIndexesList) m_A;

	[CompilerGenerated]
	private static bool m_L;

	private bool m_M;

	private static object m_A = RuntimeHelpers.GetObjectValue(new object());

	private const double m_B = 1000.0;

	private readonly IC<ActionItem> m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnAnalyze")]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("tbAnalyze")]
	private TextBlock m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkEntireWorkbook")]
	private System.Windows.Controls.CheckBox m_A;

	[AccessedThroughProperty("btnSettings")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private System.Windows.Controls.Button m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("lbxResults")]
	private System.Windows.Controls.ListBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lbxStatusItems")]
	private System.Windows.Controls.ListBox m_B;

	[AccessedThroughProperty("chkErrors")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_B;

	[AccessedThroughProperty("chkWarnings")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("chkMessages")]
	private System.Windows.Controls.CheckBox m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("chkExplanations")]
	private System.Windows.Controls.CheckBox m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFilters")]
	private System.Windows.Controls.CheckBox m_F;

	[AccessedThroughProperty("txtSearch")]
	[CompilerGenerated]
	private System.Windows.Controls.TextBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("conSettings")]
	private ContentControl m_A;

	[AccessedThroughProperty("popFilters")]
	[CompilerGenerated]
	private Popup m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("stkFilters")]
	private StackPanel m_A;

	[AccessedThroughProperty("chkFilterFormulaErrors")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_G;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFilterFormulaComplexity")]
	private System.Windows.Controls.CheckBox m_H;

	[AccessedThroughProperty("chkFilterFormulaIntegrity")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_I;

	[AccessedThroughProperty("chkFilterPerformance")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_J;

	[AccessedThroughProperty("chkFilterData")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_K;

	[AccessedThroughProperty("chkFilterHidden")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_L;

	[AccessedThroughProperty("chkFilterModelStructure")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_M;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFilterBestPractices")]
	private System.Windows.Controls.CheckBox m_N;

	[AccessedThroughProperty("chkFilterBrandCompliance")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_O;

	[AccessedThroughProperty("chkFilterPrivacySecurity")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_P;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFilterOddities")]
	private System.Windows.Controls.CheckBox Q;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFilterWorkbook")]
	private System.Windows.Controls.CheckBox R;

	private bool m_N;

	public ctpSuperFind2 thisCTP
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

	private BackgroundWorker bgw
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

	public StatusKeeper StatusKeeper
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
				switch (6)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.m_A = value;
				A(VH.A(36236));
				return;
			}
		}
	}

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

	private ObservableCollection<Observation> AllItems
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

	private Observation ActiveItem
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

	private Settings Prefs
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

	private int NavigateAfterRemoveIndex
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

	private bool CollapseAnimationRunning
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

	private List<Observation> ItemsQueuedToRemove
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

	private List<Type> DismissedFixAllTypes
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

	private bool RefocusTaskPane
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal static wpfAudit InstanceRunningAnalysis
	{
		[CompilerGenerated]
		get
		{
			return wpfAudit.m_A;
		}
		[CompilerGenerated]
		set
		{
			wpfAudit.m_A = value;
		}
	}

	public bool IsRunningAnalysis
	{
		get
		{
			return this.m_C;
		}
		set
		{
			if (!A(ref this.m_C, value, C: false, VH.A(36294)))
			{
				return;
			}
			if (this.m_C)
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
				InstanceRunningAnalysis?.A(A: true);
				InstanceRunningAnalysis = this;
			}
			else
			{
				InstanceRunningAnalysis = null;
			}
			Pane.B();
		}
	}

	public bool IsSpawnedAnalysis
	{
		get
		{
			return this.m_D;
		}
		set
		{
			A(ref this.m_D, value, C: false, VH.A(36329));
		}
	}

	public bool IsProcessingResults
	{
		get
		{
			return this.m_E;
		}
		set
		{
			if (!A(ref this.m_E, value, C: false, VH.A(36364)))
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
				JH.A();
				return;
			}
		}
	}

	public System.Windows.Visibility ExplanationsVisibility
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(36403));
		}
	}

	public string CancelSkipText
	{
		get
		{
			return this.m_A;
		}
		set
		{
			if (!A(ref this.m_A, value, C: false, VH.A(36448)))
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
				A();
				return;
			}
		}
	}

	public bool HasCancelSkipText
	{
		get
		{
			return this.m_F;
		}
		set
		{
			if (!A(ref this.m_F, value, C: false, VH.A(36477)))
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
				C();
				return;
			}
		}
	}

	public bool HasCheckExceptions
	{
		get
		{
			return this.m_G;
		}
		set
		{
			if (!A(ref this.m_G, value, C: false, VH.A(36512)))
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
				C();
				return;
			}
		}
	}

	public string CheckExceptionsText
	{
		get
		{
			return this.m_B;
		}
		set
		{
			if (!A(ref this.m_B, value, C: false, VH.A(36549)))
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
				B();
				return;
			}
		}
	}

	private string CheckExceptionsTextDetailed
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	public bool HasGeneralException
	{
		get
		{
			return this.m_H;
		}
		set
		{
			if (A(ref this.m_H, value, C: false, VH.A(36588)))
			{
				C();
			}
		}
	}

	public bool ShowsNoResultsMsg
	{
		get
		{
			return this.m_I;
		}
		set
		{
			A(ref this.m_I, value, C: false, VH.A(36627));
		}
	}

	public bool IsFullSuccessfulAnalysis
	{
		get
		{
			return this.m_J;
		}
		set
		{
			A(ref this.m_J, value, C: false, VH.A(36662));
		}
	}

	public bool ShowStandByMsg
	{
		get
		{
			return this.m_K;
		}
		set
		{
			A(ref this.m_K, value, C: false, VH.A(36711));
		}
	}

	private Range HighlightedRange
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

	internal static (bool YesNo, List<int> SheetIndexesList) RunCheckOnOpen
	{
		[CompilerGenerated]
		get
		{
			return wpfAudit.m_A;
		}
		[CompilerGenerated]
		set
		{
			wpfAudit.m_A = value;
		}
	}

	internal static bool AssumeAlreadySavedOnNextRun
	{
		[CompilerGenerated]
		get
		{
			return wpfAudit.m_L;
		}
		[CompilerGenerated]
		set
		{
			wpfAudit.m_L = value;
		}
	}

	public bool IsShowingResults
	{
		get
		{
			return this.m_M;
		}
		set
		{
			if (!A(ref this.m_M, value, C: false, VH.A(36740)))
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
				C();
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnAnalyze
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
			RoutedEventHandler value2 = StartCheckAsynch;
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

	internal virtual TextBlock tbAnalyze
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

	internal virtual System.Windows.Controls.CheckBox chkEntireWorkbook
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
			MouseButtonEventHandler value2 = ScopeMouseDown;
			System.Windows.Controls.CheckBox checkBox = this.m_A;
			if (checkBox != null)
			{
				checkBox.PreviewMouseDown -= value2;
			}
			this.m_A = value;
			checkBox = this.m_A;
			if (checkBox == null)
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
				checkBox.PreviewMouseDown += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnSettings
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
			RoutedEventHandler value2 = EditSettings;
			System.Windows.Controls.Button button = this.m_B;
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
			this.m_B = value;
			button = this.m_B;
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

	internal virtual System.Windows.Controls.Button btnCancel
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
			RoutedEventHandler value2 = CancelButtonClicked;
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
				switch (7)
				{
				case 0:
					continue;
				}
				listBox.PreviewKeyDown += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.ListBox lbxStatusItems
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

	internal virtual System.Windows.Controls.CheckBox chkErrors
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

	internal virtual System.Windows.Controls.CheckBox chkWarnings
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

	internal virtual System.Windows.Controls.CheckBox chkMessages
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
			this.m_D = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkExplanations
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
			this.m_E = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFilters
	{
		[CompilerGenerated]
		get
		{
			return this.m_F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_F = value;
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
				switch (5)
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

	internal virtual ContentControl conSettings
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

	internal virtual Popup popFilters
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

	internal virtual StackPanel stkFilters
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

	internal virtual System.Windows.Controls.CheckBox chkFilterFormulaErrors
	{
		[CompilerGenerated]
		get
		{
			return this.m_G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_G = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFilterFormulaComplexity
	{
		[CompilerGenerated]
		get
		{
			return this.m_H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_H = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFilterFormulaIntegrity
	{
		[CompilerGenerated]
		get
		{
			return this.m_I;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_I = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFilterPerformance
	{
		[CompilerGenerated]
		get
		{
			return this.m_J;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_J = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFilterData
	{
		[CompilerGenerated]
		get
		{
			return this.m_K;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_K = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFilterHidden
	{
		[CompilerGenerated]
		get
		{
			return this.m_L;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_L = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFilterModelStructure
	{
		[CompilerGenerated]
		get
		{
			return this.m_M;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_M = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFilterBestPractices
	{
		[CompilerGenerated]
		get
		{
			return this.m_N;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_N = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFilterBrandCompliance
	{
		[CompilerGenerated]
		get
		{
			return this.m_O;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_O = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFilterPrivacySecurity
	{
		[CompilerGenerated]
		get
		{
			return this.m_P;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_P = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFilterOddities
	{
		[CompilerGenerated]
		get
		{
			return Q;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			Q = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFilterWorkbook
	{
		[CompilerGenerated]
		get
		{
			return R;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			R = value;
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
	}

	public wpfAudit()
	{
		base.Loaded += ViewLoaded;
		base.Unloaded += ViewUnloaded;
		base.PreviewKeyDown += CloseFiltersPopup;
		this.m_A = 700.0;
		this.m_A = null;
		NavigateAfterRemoveIndex = -1;
		this.m_A = new IC<ActionItem>([SpecialName] () => lbxStatusItems, 1000.0);
		InitializeComponent();
		Prefs = new Settings();
		chkErrors.IsChecked = true;
		chkWarnings.IsChecked = true;
		chkMessages.IsChecked = true;
		chkFilters.IsChecked = false;
		chkExplanations.IsChecked = global::A.K.Settings.AuditShowExplanations;
		int explanationsVisibility;
		if (chkExplanations.IsChecked != true)
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
			explanationsVisibility = 2;
		}
		else
		{
			explanationsVisibility = 0;
		}
		ExplanationsVisibility = (System.Windows.Visibility)explanationsVisibility;
		E();
	}

	private void A(string A)
	{
		this.m_A?.Invoke(this, new PropertyChangedEventArgs(A));
	}

	private bool A<A>(ref A A, A B, bool C = false, [CallerMemberName] string D = null)
	{
		if (!C && object.Equals(A, B))
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
					return false;
				}
			}
		}
		A = B;
		this.A(D);
		return true;
	}

	private void A()
	{
		HasCancelSkipText = false;
	}

	private void B()
	{
		HasCheckExceptions = !modFunctionsStr.IsBlank(CheckExceptionsText);
	}

	private void C()
	{
		int isFullSuccessfulAnalysis;
		if (IsShowingResults)
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
			if (!HasCheckExceptions && !HasGeneralException)
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
				isFullSuccessfulAnalysis = ((!HasCancelSkipText) ? 1 : 0);
			}
			else
			{
				isFullSuccessfulAnalysis = 0;
			}
		}
		else
		{
			isFullSuccessfulAnalysis = 0;
		}
		IsFullSuccessfulAnalysis = (byte)isFullSuccessfulAnalysis != 0;
	}

	private void ViewLoaded(object sender, RoutedEventArgs e)
	{
		chkEntireWorkbook.Checked += ScopeChanged;
		chkEntireWorkbook.Unchecked += ScopeChanged;
		chkErrors.Checked += ToggleSeverity;
		chkErrors.Unchecked += ToggleSeverity;
		chkWarnings.Checked += ToggleSeverity;
		chkWarnings.Unchecked += ToggleSeverity;
		chkMessages.Checked += ToggleSeverity;
		chkMessages.Unchecked += ToggleSeverity;
		chkExplanations.Checked += ToggleExplanations;
		chkExplanations.Unchecked += ToggleExplanations;
		popFilters.Opened += FiltersPopupOpened;
		chkEntireWorkbook.IsChecked = true;
	}

	private void ViewUnloaded(object sender, RoutedEventArgs e)
	{
		chkEntireWorkbook.Checked -= ScopeChanged;
		chkEntireWorkbook.Unchecked -= ScopeChanged;
		chkErrors.Checked -= ToggleSeverity;
		chkErrors.Unchecked -= ToggleSeverity;
		chkWarnings.Checked -= ToggleSeverity;
		chkWarnings.Unchecked -= ToggleSeverity;
		chkMessages.Checked -= ToggleSeverity;
		chkMessages.Unchecked -= ToggleSeverity;
		chkExplanations.Checked -= ToggleExplanations;
		chkExplanations.Unchecked -= ToggleExplanations;
		popFilters.Opened -= FiltersPopupOpened;
		Prefs = null;
		bgw = null;
	}

	private void CloseFiltersPopup(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Escape || !popFilters.IsOpen)
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
			chkFilters.IsChecked = false;
			return;
		}
	}

	private void ScopeChanged(object sender, RoutedEventArgs e)
	{
		TextBlock textBlock = tbAnalyze;
		string text;
		if (chkEntireWorkbook.IsChecked != true)
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
			text = VH.A(36773);
		}
		else
		{
			text = VH.A(36820);
		}
		textBlock.Text = text;
		if (AllItems != null)
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
			AllItems.Clear();
		}
		btnAnalyze.IsHitTestVisible = true;
	}

	private void ScopeMouseDown(object sender, MouseButtonEventArgs e)
	{
		btnAnalyze.IsHitTestVisible = false;
		System.Windows.Controls.CheckBox checkBox = chkEntireWorkbook;
		bool? isChecked = chkEntireWorkbook.IsChecked;
		bool? isChecked2;
		if (!isChecked.HasValue)
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
			isChecked2 = isChecked;
		}
		else
		{
			isChecked2 = isChecked != true;
		}
		checkBox.IsChecked = isChecked2;
		e.Handled = true;
	}

	internal void D(bool A, bool B)
	{
		IsSpawnedAnalysis = B;
		chkEntireWorkbook.IsChecked = A;
		RunCheckOnOpen = default((bool, List<int>));
		System.Windows.Application.Current.Dispatcher.BeginInvoke((global::A.A)([SpecialName] () =>
		{
			Thread.Sleep(2500);
			F();
		}));
	}

	private void E()
	{
		bgw = new BackgroundWorker();
		BackgroundWorker backgroundWorker = bgw;
		backgroundWorker.WorkerSupportsCancellation = true;
		backgroundWorker.WorkerReportsProgress = false;
		backgroundWorker.DoWork += StartBackgroundAnalysis;
		backgroundWorker.RunWorkerCompleted += BackgroundAnalysisCompleted;
		_ = null;
	}

	private void StartCheckAsynch(object sender, RoutedEventArgs e)
	{
		F();
	}

	private void F()
	{
		object a = wpfAudit.m_A;
		ObjectFlowControl.CheckForSyncLockOnValueType(a);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(a, ref lockTaken);
			if (InstanceRunningAnalysis != null)
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
					Forms.WarningMessage(object.Equals(InstanceRunningAnalysis, this) ? VH.A(36952) : VH.A(36853));
					return;
				}
			}
			if (!A())
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
			if (!HB.A())
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					HB.A(object.Equals(chkEntireWorkbook.IsChecked, true));
					return;
				}
			}
			IsRunningAnalysis = true;
		}
		finally
		{
			if (lockTaken)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					Monitor.Exit(a);
					break;
				}
			}
		}
		bool flag = false;
		try
		{
			G();
			lbxResults.SelectionChanged -= ListBoxSelectionChanged;
			flag = true;
			bgw.RunWorkerAsync();
		}
		finally
		{
			if (!flag)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					IsRunningAnalysis = false;
					break;
				}
			}
		}
	}

	private bool A()
	{
		try
		{
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			if (application.Workbooks.Count == 0)
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
						C(VH.A(37015));
						return false;
					}
				}
			}
			Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application.ActiveWorkbook;
			if (!activeWorkbook.Saved)
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
				if (!AssumeAlreadySavedOnNextRun)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						Forms.WarningMessage(VH.A(37090));
						return false;
					}
				}
				activeWorkbook.Saved = true;
			}
			return true;
		}
		finally
		{
			AssumeAlreadySavedOnNextRun = false;
			Microsoft.Office.Interop.Excel.Workbook activeWorkbook = null;
			Microsoft.Office.Interop.Excel.Application application = null;
		}
	}

	private void G()
	{
		CancelSkipText = "";
		CheckExceptionsText = "";
		CheckExceptionsTextDetailed = "";
		HasGeneralException = false;
		ShowsNoResultsMsg = false;
		this.m_A.A();
		System.Windows.Forms.Application.DoEvents();
	}

	private void StartBackgroundAnalysis(object sender, DoWorkEventArgs e)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			A(ref e);
		});
	}

	private void BackgroundAnalysisCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			IsRunningAnalysis = false;
			if (!btnCancel.IsEnabled)
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
				btnCancel.IsEnabled = true;
			}
			A(e);
		});
	}

	private void A(ref DoWorkEventArgs A)
	{
		new Dictionary<int, int>();
		IsShowingResults = false;
		Icons.Cache = new IconCache();
		StatusKeeper = new StatusKeeper([SpecialName] (List<ActionItem> a) =>
		{
			this.A(a);
		});
		Settings prefs = Prefs;
		bool value = chkEntireWorkbook.IsChecked.Value;
		BackgroundWorker C = bgw;
		Analysis analysis = new Analysis(prefs, value, ref C, StatusKeeper, B);
		bgw = C;
		Analysis analysis2 = analysis;
		IsSpawnedAnalysis = false;
		this.A(analysis2);
		IsShowingResults = true;
		H();
		if (analysis2.A(A: false))
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
			A.Cancel = true;
		}
		CancelSkipText = analysis2.CancelSkipText;
		CheckExceptionsText = analysis2.CheckExceptionsText;
		CheckExceptionsTextDetailed = analysis2.CheckExceptionsTextDetailed;
		HasGeneralException = analysis2.HasGeneralException;
	}

	private void A(Analysis A)
	{
		IsProcessingResults = true;
		try
		{
			LC a = default(LC);
			LC CS_0024_003C_003E8__locals3 = new LC(a);
			CS_0024_003C_003E8__locals3.A = new Dictionary<string, double>();
			IEnumerator<IGrouping<string, Observation>> enumerator = default(IEnumerator<IGrouping<string, Observation>>);
			try
			{
				List<Observation> observations = A.Observations;
				Func<Observation, string> keySelector;
				if (_Closure_0024__.A == null)
				{
					keySelector = (_Closure_0024__.A = [SpecialName] (Observation observation2) => observation2.Title);
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
					keySelector = _Closure_0024__.A;
				}
				enumerator = observations.GroupBy(keySelector).GetEnumerator();
				while (enumerator.MoveNext())
				{
					IGrouping<string, Observation> current = enumerator.Current;
					Dictionary<string, double> dictionary = CS_0024_003C_003E8__locals3.A;
					string key = current.Key;
					Func<Observation, double> selector;
					if (_Closure_0024__.A != null)
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
					else
					{
						selector = (_Closure_0024__.A = [SpecialName] (Observation observation2) => (double)observation2.Severity);
					}
					dictionary[key] = current.Sum(selector) / (double)current.Count();
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_00d1;
					}
					continue;
					end_IL_00d1:
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
			IOrderedEnumerable<Observation> source = A.Observations.OrderByDescending([SpecialName] (Observation observation2) => CS_0024_003C_003E8__locals3.A[observation2.Title]);
			Func<Observation, string> keySelector2;
			if (_Closure_0024__.B == null)
			{
				keySelector2 = (_Closure_0024__.B = [SpecialName] (Observation observation2) => observation2.Title);
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
				keySelector2 = _Closure_0024__.B;
			}
			IOrderedEnumerable<Observation> source2 = source.ThenBy(keySelector2);
			Func<Observation, int> keySelector3;
			if (_Closure_0024__.A == null)
			{
				keySelector3 = (_Closure_0024__.A = [SpecialName] (Observation observation2) => observation2.SheetIndex);
			}
			else
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
				keySelector3 = _Closure_0024__.A;
			}
			IOrderedEnumerable<Observation> source3 = source2.ThenBy(keySelector3);
			Func<Observation, object> keySelector4;
			if (_Closure_0024__.A == null)
			{
				keySelector4 = (_Closure_0024__.A = [SpecialName] (Observation observation2) =>
				{
					if (observation2.Range == null)
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
								return 0;
							}
						}
					}
					return NewLateBinding.LateGet(observation2.Range.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41347), new object[0], null, null, null);
				});
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
				keySelector4 = _Closure_0024__.A;
			}
			IOrderedEnumerable<Observation> source4 = source3.ThenBy(keySelector4);
			Func<Observation, object> keySelector5;
			if (_Closure_0024__.B == null)
			{
				keySelector5 = (_Closure_0024__.B = [SpecialName] (Observation observation2) => (observation2.Range == null) ? ((object)0) : NewLateBinding.LateGet(observation2.Range.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41354), new object[0], null, null, null));
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
				keySelector5 = _Closure_0024__.B;
			}
			List<Observation> list = source4.ThenBy(keySelector5).ToList();
			AllItems = new ObservableCollection<Observation>();
			double D = 0.0;
			IEnumerator<IGrouping<string, Observation>> enumerator2 = default(IEnumerator<IGrouping<string, Observation>>);
			try
			{
				List<Observation> source5 = list;
				Func<Observation, string> keySelector6;
				if (_Closure_0024__.C == null)
				{
					keySelector6 = (_Closure_0024__.C = [SpecialName] (Observation observation2) => observation2.Title);
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
					keySelector6 = _Closure_0024__.C;
				}
				enumerator2 = source5.GroupBy(keySelector6).GetEnumerator();
				while (enumerator2.MoveNext())
				{
					List<Observation> a2 = enumerator2.Current.ToList();
					int[] array = new int[3]
					{
						global::A.GC.A(a2),
						global::A.GC.B(a2),
						global::A.GC.C(a2)
					};
					Observation observation = Observation.B(a2, Settings.MaxObsPerTitleTypeSeverity);
					observation.B(array[0], array[1], array[2], ref D);
					AllItems.Add(observation);
					observation.AllItems = AllItems;
					this.A(observation, array[0], array[1], array[2], ref D);
				}
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
			finally
			{
				if (enumerator2 != null)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						enumerator2.Dispose();
						break;
					}
				}
			}
		}
		finally
		{
			IsProcessingResults = false;
			List<Observation> list = null;
		}
	}

	private void A(Observation A, int B, int C, int D, ref double E)
	{
		using List<Observation>.Enumerator enumerator = A.Children.GetEnumerator();
		while (enumerator.MoveNext())
		{
			Observation current = enumerator.Current;
			current.B(B, C, D, ref E);
			this.A(current, B, C, D, ref E);
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

	private void H()
	{
		Observation.B();
		SourceCollection = CollectionViewSource.GetDefaultView(AllItems);
		SourceCollection.Filter = A;
		SourceCollection.SortDescriptions.Add(new SortDescription(VH.A(37179), ListSortDirection.Ascending));
	}

	private void A(List<ActionItem> A)
	{
		try
		{
			this.m_A.A(A.Where([SpecialName] (ActionItem a) => StatusKeeper.A(a)).ToList());
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void A(RunWorkerCompletedEventArgs A)
	{
		lbxResults.SelectionChanged += ListBoxSelectionChanged;
		ShowsNoResultsMsg = AllItems.Count == 0;
		if (AllItems.Count > 0)
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
			try
			{
				System.Windows.Forms.Application.DoEvents();
				this.A(RuntimeHelpers.GetObjectValue(MH.A.Application.ActiveWindow.SelectedSheets[1]));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)3, VH.A(37198));
	}

	private void BypassButtonClicked(object sender, RoutedEventArgs e)
	{
		if (!(sender is DependencyObject element) || !(lbxStatusItems.ContainerFromElement(element) is ListBoxItem listBoxItem))
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
			if (!(listBoxItem.Content is ActionItem actionItem))
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
				int num;
				if (!actionItem.A(CB.B, CB.F))
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							num = (actionItem.A(CB.F, CB.B) ? 1 : 0);
							return;
						}
					}
				}
				num = 1;
				return;
			}
		}
	}

	private void SkipButtonClicked(object sender, RoutedEventArgs e)
	{
		if (StatusKeeper == null)
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
			StatusKeeper.SkipIsOn = true;
			this.m_A.OverrideInteractedRecently = true;
			return;
		}
	}

	private void CancelButtonClicked(object sender, RoutedEventArgs e)
	{
		A(A: false);
	}

	internal void A(bool A = false)
	{
		BackgroundWorker backgroundWorker = bgw;
		bool? obj;
		if (backgroundWorker == null)
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
			obj = null;
		}
		else
		{
			obj = backgroundWorker.IsBusy;
		}
		if (!object.Equals(obj, true))
		{
			return;
		}
		btnCancel.IsEnabled = false;
		if (!A)
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
			if (StatusKeeper.A() > 30000)
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
				if (!StatusKeeper.A(VH.A(37227)))
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							btnCancel.IsEnabled = true;
							return;
						}
					}
				}
			}
		}
		bgw.CancelAsync();
	}

	public void ApplyFilter()
	{
		if (SourceCollection != null)
		{
			lbxResults.SelectionChanged -= ListBoxSelectionChanged;
			ActiveItem = null;
			SourceCollection = CollectionViewSource.GetDefaultView(lbxResults.ItemsSource);
			SourceCollection.Filter = A;
			SourceCollection.Refresh();
			lbxResults.SelectionChanged += ListBoxSelectionChanged;
			lbxResults.Focus();
		}
	}

	private bool A(object A)
	{
		Observation observation = (Observation)A;
		if (observation.IsGrouper)
		{
			return observation.Children.Any([SpecialName] (Observation a) => this.A((object)a));
		}
		int result;
		if (this.A(observation) && B(observation))
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
			result = (C(observation) ? 1 : 0);
		}
		else
		{
			result = 0;
		}
		return (byte)result != 0;
	}

	private bool A(Observation A)
	{
		return A.Severity switch
		{
			Severity.High => chkErrors.IsChecked.Value, 
			Severity.Medium => chkWarnings.IsChecked.Value, 
			Severity.Low => chkMessages.IsChecked.Value, 
			_ => false, 
		};
	}

	private bool B(Observation A)
	{
		return A.Category switch
		{
			Category.FormulaErrors => chkFilterFormulaErrors.IsChecked.Value, 
			Category.FormulaComplexity => chkFilterFormulaComplexity.IsChecked.Value, 
			Category.FormulaIntegrity => chkFilterFormulaIntegrity.IsChecked.Value, 
			Category.HiddenData => chkFilterHidden.IsChecked.Value, 
			Category.Performance => chkFilterPerformance.IsChecked.Value, 
			Category.PrivacySecurity => chkFilterPrivacySecurity.IsChecked.Value, 
			Category.Data => chkFilterData.IsChecked.Value, 
			Category.Workbook => chkFilterWorkbook.IsChecked.Value, 
			Category.ModelStructure => chkFilterModelStructure.IsChecked.Value, 
			Category.BestPractices => chkFilterBestPractices.IsChecked.Value, 
			Category.BrandCompliance => chkFilterBrandCompliance.IsChecked.Value, 
			Category.Oddities => chkFilterOddities.IsChecked.Value, 
			_ => false, 
		};
	}

	private bool C(Observation A)
	{
		object obj = txtSearch.Text;
		if (obj == null)
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
			obj = "";
		}
		string A2 = ((string)obj).ToLower();
		if (A2.Length == 0)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				return true;
			}
		}
		C<string, bool> c = [SpecialName] (string text) =>
		{
			object obj2 = text;
			if (obj2 == null)
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
				obj2 = "";
			}
			return ((string)obj2).ToLower().Contains(A2);
		};
		int result;
		if (!c(A.Title))
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
			if (!c(A.Subtitle))
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
				result = (c(A.Explanation) ? 1 : 0);
				goto IL_00af;
			}
		}
		result = 1;
		goto IL_00af;
		IL_00af:
		return (byte)result != 0;
	}

	private void B(bool A)
	{
		ShowStandByMsg = A;
	}

	private void ErrorsControl_MouseDoubleClick(object sender, MouseButtonEventArgs e)
	{
		try
		{
			System.Windows.Clipboard.SetText(CheckExceptionsTextDetailed);
			D(VH.A(37294));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			B(ex2.Message);
			ProjectData.ClearProjectError();
		}
	}

	private void ToggleSeverity(object sender, RoutedEventArgs e)
	{
		if (SourceCollection == null)
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
			ApplyFilter();
			if (lbxResults.SelectedIndex > -1)
			{
				try
				{
					O();
					return;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
					return;
				}
			}
			return;
		}
	}

	private void FiltersPopupOpened(object sender, EventArgs e)
	{
		popFilters.HorizontalOffset = -1.0 * stkFilters.ActualWidth / 2.0 + chkFilters.Width / 2.0;
	}

	private void FilterToggleChanged(object sender, RoutedEventArgs e)
	{
		if (SourceCollection != null)
		{
			ApplyFilter();
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
			switch (1)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (SourceCollection == null)
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
				ApplyFilter();
				return;
			}
		}
	}

	private void SearchTextBoxLostFocus(object sender, RoutedEventArgs e)
	{
		if (txtSearch.Text.Length <= 0 || SourceCollection == null)
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
			ApplyFilter();
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
			switch (3)
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
				switch (2)
				{
				case 0:
					continue;
				}
				txtSearch.SelectAll();
				if (SourceCollection == null)
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
					ApplyFilter();
					return;
				}
			}
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
					Observation observation = (Observation)lbxResults.SelectedItem;
					if (ActiveItem != null)
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
						if (ActiveItem.IsFixed)
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
							RemoveItem(ActiveItem, blnAnimate: true);
						}
					}
					ActiveItem = observation;
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
						if (!System.Windows.Input.Keyboard.IsKeyDown(Key.Up))
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
							if (!System.Windows.Input.Keyboard.IsKeyDown(Key.Right))
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
								if (!System.Windows.Input.Keyboard.IsKeyDown(Key.Left))
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
									if (!System.Windows.Input.Keyboard.IsKeyDown(Key.Next) && !System.Windows.Input.Keyboard.IsKeyDown(Key.Prior))
									{
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
											this.m_A.Tick += MarkAsRead;
											this.m_A.Start();
										}
										catch (Exception ex3)
										{
											ProjectData.SetProjectError(ex3);
											Exception ex4 = ex3;
											ProjectData.ClearProjectError();
										}
										if (!(observation is HiddenSheet) && !(observation is VeryHiddenSheet))
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
											A(observation);
										}
									}
								}
							}
						}
					}
					observation = null;
					return;
				}
				}
			}
		}
		ActiveItem = null;
	}

	private void MarkAsRead(object obj, EventArgs ev)
	{
		((DispatcherTimer)obj).Stop();
		try
		{
			if (ActiveItem.IsGrouper)
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
						return;
					}
				}
			}
			ActiveItem.FontWeight = FontWeights.Normal;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void A(Observation A)
	{
		if (A.Worksheet != null)
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
					if (ExcelAddIn1.Sheets.Visibility.A(A.Worksheet))
					{
						A.Worksheet.Activate();
						if (A.Range != null)
						{
							B(A);
						}
					}
					return;
				}
			}
		}
		if (A.Chart != null && ExcelAddIn1.Sheets.Visibility.A(A.Chart))
		{
			A.Chart.Activate();
		}
	}

	private void B(Observation A)
	{
		try
		{
			Ranges.ScrollIntoView(A.Range);
			A.Range.Select();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			C(VH.A(37357));
			I();
			F(A);
			ProjectData.ClearProjectError();
		}
	}

	private void I()
	{
		try
		{
			thisCTP.BringToFront();
			thisCTP.Focus();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		lbxResults.Focus();
	}

	private void C(Observation A)
	{
		if (lbxResults.SelectedItem == A)
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
			lbxResults.SelectedItem = A;
			return;
		}
	}

	private void ToggleExplanations(object sender, RoutedEventArgs e)
	{
		int explanationsVisibility;
		if (chkExplanations.IsChecked != true)
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
			explanationsVisibility = 2;
		}
		else
		{
			explanationsVisibility = 0;
		}
		ExplanationsVisibility = (System.Windows.Visibility)explanationsVisibility;
		global::A.K.Settings.AuditShowExplanations = chkExplanations.IsChecked.Value;
	}

	private void EditSettings(object sender, RoutedEventArgs e)
	{
		conSettings.Content = new wpfSettings(this);
		conSettings.Visibility = System.Windows.Visibility.Visible;
	}

	internal void J()
	{
		conSettings.Visibility = System.Windows.Visibility.Collapsed;
		conSettings.Content = null;
		Prefs = new Settings();
	}

	private void lbxResults_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
	{
		Key key = e.Key;
		switch (key)
		{
		case Key.Left:
		case Key.Up:
		case Key.Right:
		case Key.Down:
			if (!e.IsRepeat)
			{
				lbxResults.KeyUp += NavKeyUp;
			}
			return;
		case Key.Space:
			if (lbxResults.SelectedIndex <= -1)
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
				D((Observation)lbxResults.SelectedItem);
				return;
			}
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
			if (key != Key.Delete)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
			if (lbxResults.SelectedIndex <= -1)
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
				RemoveItemAndNavigate((Observation)lbxResults.SelectedItem);
				return;
			}
		}
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
			if (e.NewFocus is ListBoxItem || e.NewFocus is ToggleButton)
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
				if (!CollapseAnimationRunning)
				{
					System.Windows.Controls.ListBox listBox = lbxResults;
					if (listBox.SelectedIndex > -1 && ((Observation)listBox.SelectedItem).IsFixed)
					{
						RemoveItem((Observation)listBox.SelectedItem, blnAnimate: true);
					}
					listBox.SelectedIndex = -1;
					listBox = null;
				}
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void RemoveButtonCicked(object sender, RoutedEventArgs e)
	{
		RemoveItemAndNavigate((Observation)((System.Windows.Controls.Button)sender).DataContext);
	}

	public void RemoveItemAndNavigate(Observation err)
	{
		NavigateAfterRemoveIndex = lbxResults.SelectedIndex;
		RemoveItem(err, blnAnimate: true);
	}

	public void RemoveItem(Observation itm, bool blnAnimate)
	{
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			ItemsQueuedToRemove = new List<Observation>();
		}
		ItemsQueuedToRemove.Add(itm);
		if (blnAnimate)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
				{
					ListBoxItem listBoxItem;
					try
					{
						listBoxItem = (ListBoxItem)lbxResults.ItemContainerGenerator.ContainerFromItem(itm);
						if (listBoxItem != null)
						{
							A(listBoxItem);
						}
						else
						{
							K();
						}
					}
					catch (NullReferenceException ex)
					{
						ProjectData.SetProjectError(ex);
						NullReferenceException ex2 = ex;
						clsReporting.LogException((Exception)ex2);
						ProjectData.ClearProjectError();
					}
					listBoxItem = null;
					return;
				}
				}
			}
		}
		K();
	}

	private void FixButtonClicked(object sender, RoutedEventArgs e)
	{
		D((Observation)((System.Windows.Controls.Button)sender).DataContext);
	}

	private void D(Observation A)
	{
		C(A);
		E(A);
	}

	private void E(Observation A)
	{
		if (A.IsFixed)
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
			this.A(0);
			return;
		}
	}

	private void ShowFixOptions(object sender, RoutedEventArgs e)
	{
		Observation a = (Observation)((ToggleButton)sender).DataContext;
		C(a);
		A(a, (ToggleButton)sender, C: true);
		a = null;
	}

	private void A(Observation A, ToggleButton B, bool C)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Expected O, but got Unknown
		//IL_002c: Expected O, but got Unknown
		RefocusTaskPane = C;
		wpfFixMenu val = new wpfFixMenu(B, A.DisplayText, (Visual)this, false);
		((System.Windows.Window)val).Closed += PopupMenuClosed;
		((System.Windows.Window)val).Show();
		_ = null;
	}

	private void PopupMenuClosed(object sender, EventArgs e)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		wpfFixMenu val = (wpfFixMenu)sender;
		if (val.Index > -1)
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
			A(val.Index);
		}
		val = null;
		if (!RefocusTaskPane)
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
			O();
			return;
		}
	}

	private void A(int A)
	{
		if (ActiveItem == null)
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
			this.A(ActiveItem, A);
			return;
		}
	}

	private void A(Observation A, int B)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		XlCalculation calculation = application.Calculation;
		bool flag = false;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		if (!(A is CalcModeManual))
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
			application.Calculation = XlCalculation.xlCalculationManual;
		}
		try
		{
			this.B(A, B);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			this.B(VH.A(37390) + ex2.Message);
			flag = true;
			ProjectData.ClearProjectError();
		}
		if (A.CanFixMultiple)
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
			if (!flag)
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
				if (DismissedFixAllTypes == null)
				{
					DismissedFixAllTypes = new List<Type>();
				}
				if (!DismissedFixAllTypes.Contains(A.GetType()))
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
						List<Observation> list = wpfAudit.A(AllItems, A);
						int count = list.Count;
						if (count > 0)
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
							if (Forms.YesNoMessage2(VH.A(37439) + count + VH.A(37506), (YesNoDefault)1) == DialogResult.Yes)
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
									using (List<Observation>.Enumerator enumerator = list.GetEnumerator())
									{
										while (enumerator.MoveNext())
										{
											Observation current = enumerator.Current;
											this.B(current, B);
											RemoveItem(current, blnAnimate: false);
										}
										while (true)
										{
											switch (5)
											{
											case 0:
												break;
											default:
												goto end_IL_018c;
											}
											continue;
											end_IL_018c:
											break;
										}
									}
									RemoveItemAndNavigate(A);
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									this.B(VH.A(37543) + ex4.Message);
									ProjectData.ClearProjectError();
								}
								wpfAudit.A(AllItems).ForEach([SpecialName] (Observation itm) =>
								{
									RemoveItem(itm, blnAnimate: false);
								});
							}
							else
							{
								DismissedFixAllTypes.Add(Information.Err().GetType());
							}
							I();
						}
						list = null;
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						this.B(ex6.Message);
						ProjectData.ClearProjectError();
					}
				}
			}
		}
		if (!(A is CalcModeManual))
		{
			application.Calculation = calculation;
		}
		application.EnableEvents = true;
		application.ScreenUpdating = true;
		application = null;
	}

	private static List<Observation> A(IEnumerable<Observation> A, Observation B)
	{
		List<Observation> list = new List<Observation>();
		IEnumerator<Observation> enumerator = default(IEnumerator<Observation>);
		try
		{
			enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Observation current = enumerator.Current;
				if (current == B)
				{
					continue;
				}
				if ((object)current.GetType() == B.GetType())
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
					list.Add(current);
				}
				else
				{
					if (!(current is FC))
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
					if (current.Children == null)
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
						break;
					}
					list.AddRange(wpfAudit.A(current.Children, B));
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				return list;
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
	}

	private static List<Observation> A(IEnumerable<Observation> A)
	{
		List<Observation> list = new List<Observation>();
		using IEnumerator<Observation> enumerator = A.Where([SpecialName] (Observation observation) => observation is FC).GetEnumerator();
		while (enumerator.MoveNext())
		{
			Observation current = enumerator.Current;
			List<Observation> list2 = current.Children;
			if (list2 == null)
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
				list2 = new List<Observation>();
			}
			List<Observation> list3 = list2;
			if (list3.Count == 0)
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
				list.Add(current);
			}
			else
			{
				list.AddRange(wpfAudit.A(list3));
			}
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			return list;
		}
	}

	private void B(Observation A, int B)
	{
		A.FixAction();
		A.FixAction(B);
		A.IsFixed = true;
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
		K();
		CollapseAnimationRunning = false;
	}

	private void K()
	{
		L();
		F(ItemsQueuedToRemove[0]);
		ItemsQueuedToRemove.RemoveAt(0);
		if (NavigateAfterRemoveIndex > -1)
		{
			lbxResults.SelectedIndex = NavigateAfterRemoveIndex;
			O();
			NavigateAfterRemoveIndex = -1;
		}
	}

	private void L()
	{
		NC a = default(NC);
		NC CS_0024_003C_003E8__locals3 = new NC(a);
		CS_0024_003C_003E8__locals3.A = ItemsQueuedToRemove[0];
		checked
		{
			foreach (Observation item in AllItems.Where([SpecialName] (Observation A) => object.Equals(A.Title, CS_0024_003C_003E8__locals3.A.Title)))
			{
				Severity severity = CS_0024_003C_003E8__locals3.A.Severity;
				if (severity != Severity.Medium)
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
					if (severity == Severity.High)
					{
						item.ErrorsCount--;
					}
					else
					{
						item.MessagesCount--;
					}
				}
				else
				{
					item.WarningsCount--;
				}
			}
		}
	}

	private void F(Observation A)
	{
		Observation parent = A.Parent;
		if (parent == null)
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
		}
		else
		{
			parent.A(A);
		}
		AllItems.Remove(A);
	}

	private void A(object A)
	{
		OC a = default(OC);
		OC CS_0024_003C_003E8__locals2 = new OC(a);
		CS_0024_003C_003E8__locals2.A = A;
		if (lbxResults.IsKeyboardFocusWithin)
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
			if (lbxResults.IsKeyboardFocused)
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
				try
				{
					int index = SourceCollection.OfType<Observation>().ToList().IndexOf(SourceCollection.OfType<Observation>().First([SpecialName] (Observation observation) => Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(observation.Sheet, null, VH.A(48135), new object[0], null, null, null), NewLateBinding.LateGet(CS_0024_003C_003E8__locals2.A, null, VH.A(48135), new object[0], null, null, null), TextCompare: false)));
					ListBoxItem obj = (ListBoxItem)lbxResults.ItemContainerGenerator.ContainerFromItem(RuntimeHelpers.GetObjectValue(lbxResults.Items[index]));
					ScrollViewer scrollViewer = Forms.GetScrollViewer((DependencyObject)lbxResults) as ScrollViewer;
					Visual visual = obj;
					GroupItem groupItem = null;
					while (true)
					{
						if (visual != null)
						{
							visual = VisualTreeHelper.GetParent(visual) as Visual;
							if (visual == null)
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
								break;
							}
							if (!(visual is GroupItem))
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
							groupItem = visual as GroupItem;
							Expander obj2 = (Expander)this.A(groupItem, typeof(Expander));
							ItemsPresenter relativeTo = (ItemsPresenter)scrollViewer.Content;
							scrollViewer.ScrollToVerticalOffset(obj2.TranslatePoint(default(System.Windows.Point), relativeTo).Y);
							relativeTo = null;
							System.Windows.Point point = default(System.Windows.Point);
							break;
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
						break;
					}
					scrollViewer = null;
					visual = null;
					groupItem = null;
					return;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
					return;
				}
			}
		}
	}

	private Visual A(Visual A, Type B)
	{
		if (A == null)
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
					return null;
				}
			}
		}
		if (A.GetType() == B)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					return A;
				}
			}
		}
		Visual visual = null;
		if (A is FrameworkElement)
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
			(A as FrameworkElement).ApplyTemplate();
		}
		checked
		{
			int num = VisualTreeHelper.GetChildrenCount(A) - 1;
			for (int i = 0; i <= num; i++)
			{
				Visual a = VisualTreeHelper.GetChild(A, i) as Visual;
				visual = this.A(a, B);
				if (visual != null)
				{
					break;
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_008b;
					}
					continue;
					end_IL_008b:
					break;
				}
			}
			return visual;
		}
	}

	private void M()
	{
		throw new NotImplementedException();
	}

	private void N()
	{
		if (HighlightedRange == null)
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
			Highlight.B(HighlightedRange);
			HighlightedRange = null;
			return;
		}
	}

	private void O()
	{
		Forms.FocusListBoxItem(lbxResults, true);
	}

	private void B(string A)
	{
		Forms.ErrorMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void C(string A)
	{
		Forms.WarningMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void D(string A)
	{
		Forms.InfoMessage(System.Windows.Window.GetWindow(this), A);
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (this.m_N)
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
			this.m_N = true;
			Uri resourceLocator = new Uri(VH.A(37658), UriKind.Relative);
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
					btnAnalyze = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			tbAnalyze = (TextBlock)target;
			return;
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkEntireWorkbook = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			btnSettings = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 6)
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
		if (connectionId == 7)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					((ContentControl)target).MouseDoubleClick += ErrorsControl_MouseDoubleClick;
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
					lbxResults = (System.Windows.Controls.ListBox)target;
					return;
				}
			}
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
					lbxStatusItems = (System.Windows.Controls.ListBox)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			chkErrors = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkWarnings = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			chkMessages = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 20)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkExplanations = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 21)
		{
			chkFilters = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 22)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					txtSearch = (System.Windows.Controls.TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 23)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					conSettings = (ContentControl)target;
					return;
				}
			}
		}
		if (connectionId == 24)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					popFilters = (Popup)target;
					return;
				}
			}
		}
		if (connectionId == 25)
		{
			stkFilters = (StackPanel)target;
			return;
		}
		if (connectionId == 26)
		{
			chkFilterFormulaErrors = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 27)
		{
			chkFilterFormulaComplexity = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 28)
		{
			chkFilterFormulaIntegrity = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 29)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkFilterPerformance = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 30)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkFilterData = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 31)
		{
			chkFilterHidden = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 32)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkFilterModelStructure = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 33)
		{
			chkFilterBestPractices = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 34)
		{
			chkFilterBrandCompliance = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 35)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkFilterPrivacySecurity = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 36:
			chkFilterOddities = (System.Windows.Controls.CheckBox)target;
			break;
		case 37:
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				chkFilterWorkbook = (System.Windows.Controls.CheckBox)target;
				return;
			}
		default:
			this.m_N = true;
			break;
		}
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
		if (connectionId == 1)
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
			eventSetter.Event = ToggleButton.CheckedEvent;
			eventSetter.Handler = new RoutedEventHandler(FilterToggleChanged);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = ToggleButton.UncheckedEvent;
			eventSetter.Handler = new RoutedEventHandler(FilterToggleChanged);
			((System.Windows.Style)target).Setters.Add(eventSetter);
		}
		if (connectionId == 9)
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
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = UIElement.LostKeyboardFocusEvent;
			eventSetter.Handler = new KeyboardFocusChangedEventHandler(ListBoxItemLostKeyboardFocus);
			((System.Windows.Style)target).Setters.Add(eventSetter);
		}
		if (connectionId == 10)
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
			((System.Windows.Controls.Button)target).Click += RemoveButtonCicked;
		}
		if (connectionId == 11)
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
			((System.Windows.Controls.Button)target).Click += FixButtonClicked;
		}
		if (connectionId == 12)
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
			((ToggleButton)target).Checked += ShowFixOptions;
		}
		if (connectionId == 14)
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
			((System.Windows.Controls.Button)target).Click += BypassButtonClicked;
		}
		if (connectionId == 15)
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
			((System.Windows.Controls.Button)target).Click += SkipButtonClicked;
		}
		if (connectionId != 16)
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
			((System.Windows.Controls.Button)target).Click += BypassButtonClicked;
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
	private System.Windows.Controls.ListBox A()
	{
		return lbxStatusItems;
	}

	[SpecialName]
	[CompilerGenerated]
	private void P()
	{
		Thread.Sleep(2500);
		F();
	}

	[SpecialName]
	[CompilerGenerated]
	private void B(List<ActionItem> A)
	{
		this.A(A);
	}

	[SpecialName]
	[CompilerGenerated]
	private bool A(ActionItem A)
	{
		return StatusKeeper.A(A);
	}

	[SpecialName]
	[CompilerGenerated]
	private bool D(Observation A)
	{
		return this.A((object)A);
	}

	[SpecialName]
	[CompilerGenerated]
	private void G(Observation A)
	{
		RemoveItem(A, blnAnimate: false);
	}
}
