using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts.MoveDataLabels;

[DesignerGenerated]
public sealed class wpfPane : System.Windows.Controls.UserControl, INotifyPropertyChanged, IComponentConnector
{
	[CompilerGenerated]
	internal sealed class ID
	{
		public Series A;

		public int A;

		[SpecialName]
		internal bool A(SeriesItem A)
		{
			if (A.A.PlotOrder == this.A.PlotOrder)
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
						return A.A.AxisGroup == this.A.AxisGroup;
					}
				}
			}
			return false;
		}

		[SpecialName]
		internal bool A(KeyValuePair<int, int> A)
		{
			return A.Value == this.A;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private bool m_A;

	private bool m_B;

	private bool m_C;

	private int m_A;

	private int m_B;

	private int m_C;

	[AccessedThroughProperty("timer")]
	[CompilerGenerated]
	private System.Windows.Forms.Timer m_A;

	[CompilerGenerated]
	private Chart m_A;

	private ObservableCollection<SeriesItem> m_A;

	private ObservableCollection<LabelsItem> m_A;

	private ObservableCollection<Tuple<string, XlDataLabelPosition>> m_A;

	private ObservableCollection<Tuple<string, XlOrientation>> m_A;

	[CompilerGenerated]
	private Dictionary<int, int> m_A;

	[CompilerGenerated]
	private Action<DataLabel, int> m_A;

	[CompilerGenerated]
	private Dictionary<DataLabel, System.Windows.Point> m_A;

	private string m_A;

	private Visibility m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("scroller")]
	private ScrollViewer m_A;

	[AccessedThroughProperty("stkContents")]
	[CompilerGenerated]
	private StackPanel m_A;

	[AccessedThroughProperty("cbxSeries")]
	[CompilerGenerated]
	private System.Windows.Controls.ComboBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxLabel")]
	private System.Windows.Controls.ComboBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxPosition")]
	private System.Windows.Controls.ComboBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnUp")]
	private System.Windows.Controls.Button m_A;

	[AccessedThroughProperty("btnLeft")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("txtStep")]
	private System.Windows.Controls.TextBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnRight")]
	private System.Windows.Controls.Button m_C;

	[AccessedThroughProperty("btnDown")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_D;

	[AccessedThroughProperty("cbxRotation")]
	[CompilerGenerated]
	private System.Windows.Controls.ComboBox m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("btnReset")]
	private System.Windows.Controls.Button m_E;

	[AccessedThroughProperty("btnClose")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_F;

	[AccessedThroughProperty("btnDismiss")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_G;

	private bool m_D;

	private virtual System.Windows.Forms.Timer timer
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
			EventHandler value2 = MoveRepeat;
			System.Windows.Forms.Timer timer = this.m_A;
			if (timer != null)
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
				timer.Tick -= value2;
			}
			this.m_A = value;
			timer = this.m_A;
			if (timer == null)
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
				timer.Tick += value2;
				return;
			}
		}
	}

	private Chart Chart
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

	public ObservableCollection<SeriesItem> SeriesCollection
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(56419));
		}
	}

	public ObservableCollection<LabelsItem> LabelsCollection
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(56452));
		}
	}

	public ObservableCollection<Tuple<string, XlDataLabelPosition>> LabelPositions
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(56485));
		}
	}

	public ObservableCollection<Tuple<string, XlOrientation>> LabelRotations
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(56514));
		}
	}

	private Dictionary<int, int> LabelDictionary
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

	private Action<DataLabel, int> MoveAction
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

	private Dictionary<DataLabel, System.Windows.Point> MovedLabels
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

	public string ErrorText
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(53064));
			int errorVisibility;
			if (value.Length <= 0)
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
				errorVisibility = 2;
			}
			else
			{
				errorVisibility = 0;
			}
			ErrorVisibility = (Visibility)errorVisibility;
		}
	}

	public Visibility ErrorVisibility
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(53083));
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

	internal virtual StackPanel stkContents
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

	internal virtual System.Windows.Controls.ComboBox cbxSeries
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

	internal virtual System.Windows.Controls.ComboBox cbxLabel
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

	internal virtual System.Windows.Controls.ComboBox cbxPosition
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

	internal virtual System.Windows.Controls.Button btnUp
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
			MouseButtonEventHandler value2 = ArrowMouseDown;
			MouseButtonEventHandler value3 = ArrowMouseUp;
			System.Windows.Input.KeyEventHandler value4 = ArrowKeyDown;
			System.Windows.Input.KeyEventHandler value5 = ArrowKeyUp;
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
				button.PreviewMouseDown -= value2;
				button.PreviewMouseUp -= value3;
				button.PreviewKeyDown -= value4;
				button.PreviewKeyUp -= value5;
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
				button.PreviewMouseDown += value2;
				button.PreviewMouseUp += value3;
				button.PreviewKeyDown += value4;
				button.PreviewKeyUp += value5;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnLeft
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
			MouseButtonEventHandler value2 = ArrowMouseDown;
			MouseButtonEventHandler value3 = ArrowMouseUp;
			System.Windows.Input.KeyEventHandler value4 = ArrowKeyDown;
			System.Windows.Input.KeyEventHandler value5 = ArrowKeyUp;
			System.Windows.Controls.Button button = this.m_B;
			if (button != null)
			{
				button.PreviewMouseDown -= value2;
				button.PreviewMouseUp -= value3;
				button.PreviewKeyDown -= value4;
				button.PreviewKeyUp -= value5;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				button.PreviewMouseDown += value2;
				button.PreviewMouseUp += value3;
				button.PreviewKeyDown += value4;
				button.PreviewKeyUp += value5;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.TextBox txtStep
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

	internal virtual System.Windows.Controls.Button btnRight
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
			MouseButtonEventHandler value2 = ArrowMouseDown;
			MouseButtonEventHandler value3 = ArrowMouseUp;
			System.Windows.Input.KeyEventHandler value4 = ArrowKeyDown;
			System.Windows.Input.KeyEventHandler value5 = ArrowKeyUp;
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
				button.PreviewMouseDown -= value2;
				button.PreviewMouseUp -= value3;
				button.PreviewKeyDown -= value4;
				button.PreviewKeyUp -= value5;
			}
			this.m_C = value;
			button = this.m_C;
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
				button.PreviewMouseDown += value2;
				button.PreviewMouseUp += value3;
				button.PreviewKeyDown += value4;
				button.PreviewKeyUp += value5;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnDown
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
			MouseButtonEventHandler value2 = ArrowMouseDown;
			MouseButtonEventHandler value3 = ArrowMouseUp;
			System.Windows.Input.KeyEventHandler value4 = ArrowKeyDown;
			System.Windows.Input.KeyEventHandler value5 = ArrowKeyUp;
			System.Windows.Controls.Button button = this.m_D;
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
				button.PreviewMouseDown -= value2;
				button.PreviewMouseUp -= value3;
				button.PreviewKeyDown -= value4;
				button.PreviewKeyUp -= value5;
			}
			this.m_D = value;
			button = this.m_D;
			if (button != null)
			{
				button.PreviewMouseDown += value2;
				button.PreviewMouseUp += value3;
				button.PreviewKeyDown += value4;
				button.PreviewKeyUp += value5;
			}
		}
	}

	internal virtual System.Windows.Controls.ComboBox cbxRotation
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

	internal virtual System.Windows.Controls.Button btnReset
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
			RoutedEventHandler value2 = btnReset_Click;
			System.Windows.Controls.Button button = this.m_E;
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
			this.m_E = value;
			button = this.m_E;
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

	internal virtual System.Windows.Controls.Button btnClose
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
			RoutedEventHandler value2 = btnClose_Click;
			System.Windows.Controls.Button button = this.m_F;
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
			this.m_F = value;
			button = this.m_F;
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

	internal virtual System.Windows.Controls.Button btnDismiss
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
			RoutedEventHandler value2 = DismissError;
			System.Windows.Controls.Button button = this.m_G;
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
			this.m_G = value;
			button = this.m_G;
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

	public wpfPane()
	{
		base.Unloaded += wpfPane_Unloaded;
		base.KeyDown += HandleKeyPresses;
		base.PreviewKeyDown += wpfMoveDataLabels_PreviewKeyDown;
		timer = new System.Windows.Forms.Timer();
		this.m_A = "";
		this.m_A = Visibility.Collapsed;
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
			switch (6)
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

	private void wpfPane_Unloaded(object sender, RoutedEventArgs e)
	{
	}

	private void PaneSizeChanged(object sender, SizeChangedEventArgs e)
	{
		Panes.PaneSizeChanged(scroller, e);
	}

	private void ContentsHeightChanged(object sender, SizeChangedEventArgs e)
	{
		if (e.HeightChanged)
		{
			Panes.UpdateScrollViewerPadding(scroller);
		}
	}

	public void ShowPane()
	{
		this.m_A = false;
		this.m_A = -1;
		this.m_B = -1;
		MovedLabels = new Dictionary<DataLabel, System.Windows.Point>();
		LabelPositions = new ObservableCollection<Tuple<string, XlDataLabelPosition>>();
		ObservableCollection<Tuple<string, XlDataLabelPosition>> labelPositions = LabelPositions;
		labelPositions.Add(new Tuple<string, XlDataLabelPosition>(VH.A(56543), XlDataLabelPosition.xlLabelPositionCenter));
		labelPositions.Add(new Tuple<string, XlDataLabelPosition>(VH.A(56560), XlDataLabelPosition.xlLabelPositionAbove));
		labelPositions.Add(new Tuple<string, XlDataLabelPosition>(VH.A(56571), XlDataLabelPosition.xlLabelPositionBelow));
		labelPositions.Add(new Tuple<string, XlDataLabelPosition>(VH.A(56582), XlDataLabelPosition.xlLabelPositionLeft));
		labelPositions.Add(new Tuple<string, XlDataLabelPosition>(VH.A(56591), XlDataLabelPosition.xlLabelPositionRight));
		labelPositions.Add(new Tuple<string, XlDataLabelPosition>(VH.A(56602), XlDataLabelPosition.xlLabelPositionInsideBase));
		labelPositions.Add(new Tuple<string, XlDataLabelPosition>(VH.A(56625), XlDataLabelPosition.xlLabelPositionInsideEnd));
		labelPositions.Add(new Tuple<string, XlDataLabelPosition>(VH.A(56646), XlDataLabelPosition.xlLabelPositionOutsideEnd));
		_ = null;
		LabelRotations = new ObservableCollection<Tuple<string, XlOrientation>>();
		ObservableCollection<Tuple<string, XlOrientation>> labelRotations = LabelRotations;
		labelRotations.Add(new Tuple<string, XlOrientation>(VH.A(56669), XlOrientation.xlHorizontal));
		labelRotations.Add(new Tuple<string, XlOrientation>(VH.A(56690), XlOrientation.xlDownward));
		labelRotations.Add(new Tuple<string, XlOrientation>(VH.A(56711), XlOrientation.xlUpward));
		labelRotations.Add(new Tuple<string, XlOrientation>(VH.A(56734), XlOrientation.xlVertical));
		_ = null;
		base.SizeChanged -= PaneSizeChanged;
		base.SizeChanged += PaneSizeChanged;
		stkContents.SizeChanged -= ContentsHeightChanged;
		stkContents.SizeChanged += ContentsHeightChanged;
		A();
	}

	public void HidePane()
	{
		if (int.TryParse(txtStep.Text, out var result))
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
			global::A.K.Settings.MoveDataLabelsStep = result;
		}
		N();
		base.SizeChanged -= PaneSizeChanged;
		stkContents.SizeChanged -= ContentsHeightChanged;
		if (!this.m_A)
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
			A();
		}
		Chart = null;
		SeriesCollection = null;
		LabelsCollection = null;
		LabelDictionary = null;
		LabelPositions = null;
		LabelRotations = null;
		MovedLabels = null;
		MoveAction = null;
		timer = null;
	}

	private void A()
	{
		this.m_A = true;
		Pane.B();
	}

	private void btnClose_Click(object sender, RoutedEventArgs e)
	{
		A();
	}

	private void DismissError(object sender, RoutedEventArgs e)
	{
		A();
	}

	private void HandleKeyPresses(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Return)
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
			if (e.Key != Key.Escape)
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
				break;
			}
		}
		A();
	}

	private void B()
	{
		try
		{
			new ComAwareEventInfo(typeof(ChartEvents_Event), VH.A(53225)).RemoveEventHandler(Chart, new ChartEvents_DeactivateEventHandler(B));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (!A())
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
					A();
					return;
				}
			}
		}
		C();
	}

	private void C()
	{
		MH.A.Application.CommandBars.ReleaseFocus();
	}

	private void B(string A)
	{
		ErrorText = A;
	}

	private void D()
	{
		ErrorText = "";
	}

	private bool A()
	{
		Chart = Helpers.SelectedChart();
		if (Chart != null)
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
					new ComAwareEventInfo(typeof(ChartEvents_Event), VH.A(53225)).RemoveEventHandler(Chart, new ChartEvents_DeactivateEventHandler(B));
					new ComAwareEventInfo(typeof(ChartEvents_Event), VH.A(53225)).AddEventHandler(Chart, new ChartEvents_DeactivateEventHandler(B));
					D();
					E();
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(56371));
					return true;
				}
			}
		}
		B(VH.A(53749));
		return false;
	}

	private void E()
	{
		Visibility b = Visibility.Visible;
		N();
		cbxSeries.SelectionChanged += SeriesSelectionChanged;
		this.m_C = false;
		this.m_B = false;
		Series A;
		try
		{
			txtStep.Text = Conversions.ToString(global::A.K.Settings.MoveDataLabelsStep);
			int count = ((SeriesCollection)Chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Count;
			if (count == 1)
			{
				XlChartType chartType = Chart.ChartType;
				if (chartType <= XlChartType.xl3DPie)
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
					if (chartType != XlChartType.xlDoughnut)
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
						if (chartType != XlChartType.xl3DPie)
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
							goto IL_0113;
						}
					}
				}
				else if (chartType != XlChartType.xlPie)
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
					if ((uint)(chartType - 68) > 3u)
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
						if (chartType != XlChartType.xlDoughnutExploded)
						{
							goto IL_0113;
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
				this.m_B = true;
				b = Visibility.Collapsed;
			}
			goto IL_0113;
			IL_0113:
			this.A(count, b);
			object objectValue = RuntimeHelpers.GetObjectValue(MH.A.Application.Selection);
			if (objectValue is Microsoft.Office.Interop.Excel.DataLabels)
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
				A = (Series)((Microsoft.Office.Interop.Excel.DataLabels)objectValue).Parent;
			}
			else if (objectValue is DataLabel)
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
				A = (Series)((Microsoft.Office.Interop.Excel.Point)((DataLabel)objectValue).Parent).Parent;
			}
			else if (objectValue is Series)
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
				A = (Series)objectValue;
			}
			else
			{
				A = (Series)Chart.SeriesCollection(1);
				try
				{
					NewLateBinding.LateCall(A.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(51162), new object[0], null, null, null, IgnoreReturn: true);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			try
			{
				cbxSeries.SelectedItem = SeriesCollection.Where([SpecialName] (SeriesItem seriesItem) =>
				{
					if (seriesItem.A.PlotOrder == A.PlotOrder)
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
								return seriesItem.A.AxisGroup == A.AxisGroup;
							}
						}
					}
					return false;
				}).ElementAtOrDefault(0);
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			this.m_C = true;
			if (objectValue is DataLabel)
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
				string input = Conversions.ToString(Chart.Application.ExecuteExcel4Macro(VH.A(56749)));
				int A2 = Conversions.ToInteger(Regex.Match(input, VH.A(56772)).Groups[1].Value);
				List<int> list = (from keyValuePair in LabelDictionary
					where keyValuePair.Value == A2
					select keyValuePair.Key).ToList();
				cbxLabel.SelectedIndex = list[0];
				this.A((DataLabel)objectValue);
			}
			else if (A.HasDataLabels)
			{
				this.A((Microsoft.Office.Interop.Excel.DataLabels)A.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)));
			}
			cbxLabel.SelectionChanged += LabelSelectionChanged;
			L();
			btnReset.IsEnabled = cbxPosition.SelectedIndex > -1;
			cbxSeries.IsEnabled = count > 1;
			txtStep.Focus();
			txtStep.SelectAll();
		}
		finally
		{
			A = null;
			List<int> list = null;
			object objectValue = null;
		}
	}

	private void A(int A, Visibility B)
	{
		SeriesCollection = new ObservableCollection<SeriesItem>();
		for (int i = 1; i <= A; i = checked(i + 1))
		{
			Series series = (Series)Chart.SeriesCollection(i);
			if (this.A(series))
			{
				SeriesCollection.Add(new SeriesItem(series, i, B));
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
			Series series = null;
			return;
		}
	}

	private void SeriesSelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (((System.Windows.Controls.ComboBox)sender).SelectedIndex == -1)
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
			LabelsCollection = new ObservableCollection<LabelsItem>();
			LabelsCollection.Add(new LabelsItem(VH.A(56789), new SolidColorBrush(Colors.Black), Visibility.Collapsed));
			Microsoft.Office.Interop.Excel.DataLabels dataLabels = (Microsoft.Office.Interop.Excel.DataLabels)A().DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
			LabelDictionary = new Dictionary<int, int>();
			int count = dataLabels.Count;
			for (int i = 1; i <= count; i = checked(i + 1))
			{
				try
				{
					if (dataLabels.Item(i).Text.Length <= 0)
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
						LabelDictionary.Add(LabelsCollection.Count, i);
						string strLabel = i + VH.A(56830) + dataLabels.Item(i).Text;
						if (this.m_B)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								LabelsCollection.Add(new LabelsItem(strLabel, ColorTile.A(((Microsoft.Office.Interop.Excel.Point)A().Points(i)).Format.Fill), Visibility.Visible));
								break;
							}
						}
						else
						{
							LabelsCollection.Add(new LabelsItem(strLabel, new SolidColorBrush(Colors.Black), Visibility.Collapsed));
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
					continue;
				}
				cbxLabel.SelectedIndex = 0;
				if (this.m_C)
				{
					M();
					F();
					L();
				}
				dataLabels = null;
				return;
			}
		}
	}

	private void LabelSelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (((System.Windows.Controls.ComboBox)sender).SelectedIndex == -1)
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
			M();
			if (B())
			{
				F();
			}
			else
			{
				DataLabel dataLabel = A();
				cbxLabel.SelectionChanged -= LabelSelectionChanged;
				try
				{
					dataLabel.Select();
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				cbxLabel.SelectionChanged += LabelSelectionChanged;
				A(dataLabel);
				dataLabel = null;
			}
			L();
			cbxLabel.Focus();
			return;
		}
	}

	private void PositionSelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
		bool flag = comboBox.SelectedIndex > -1;
		if (flag)
		{
			XlDataLabelPosition position = (XlDataLabelPosition)Conversions.ToInteger(comboBox.SelectedValue);
			try
			{
				if (B())
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
					((Microsoft.Office.Interop.Excel.DataLabels)A().DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Position = position;
				}
				else
				{
					A().Position = position;
				}
				this.m_A = comboBox.SelectedIndex;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.WarningMessage(VH.A(56837));
				I();
				comboBox.SelectedIndex = this.m_A;
				H();
				ProjectData.ClearProjectError();
			}
		}
		btnReset.IsEnabled = flag;
		comboBox = null;
	}

	private void RotationSelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
		bool flag = comboBox.SelectedIndex > -1;
		if (flag)
		{
			XlOrientation xlOrientation = (XlOrientation)Conversions.ToInteger(comboBox.SelectedValue);
			try
			{
				if (B())
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
					((Microsoft.Office.Interop.Excel.DataLabels)A().DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Orientation = xlOrientation;
				}
				else
				{
					A().Orientation = xlOrientation;
				}
				this.m_B = comboBox.SelectedIndex;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.WarningMessage(VH.A(56924));
				K();
				comboBox.SelectedIndex = this.m_B;
				J();
				ProjectData.ClearProjectError();
			}
		}
		btnReset.IsEnabled = flag;
		comboBox = null;
	}

	private bool B()
	{
		return cbxLabel.SelectedIndex == 0;
	}

	private Series A()
	{
		Series result;
		try
		{
			result = ((SeriesItem)cbxSeries.SelectedItem).A;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = null;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private DataLabel A()
	{
		if (cbxLabel.SelectedIndex > -1)
		{
			return (DataLabel)A().DataLabels(LabelDictionary[cbxLabel.SelectedIndex]);
		}
		return null;
	}

	private void F()
	{
		Microsoft.Office.Interop.Excel.DataLabels dataLabels = (Microsoft.Office.Interop.Excel.DataLabels)A().DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
		dataLabels.Select();
		A(dataLabels);
		dataLabels = null;
	}

	private void A(DataLabel A)
	{
		this.A(A.Position);
		this.A((XlOrientation)Conversions.ToInteger(A.Orientation));
	}

	private void A(Microsoft.Office.Interop.Excel.DataLabels A)
	{
		this.A(A.Position);
		this.A((XlOrientation)Conversions.ToInteger(A.Orientation));
	}

	private void A(XlDataLabelPosition A)
	{
		using (IEnumerator<Tuple<string, XlDataLabelPosition>> enumerator = LabelPositions.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				Tuple<string, XlDataLabelPosition> current = enumerator.Current;
				if (current.Item2 == A)
				{
					cbxPosition.SelectedItem = current;
					return;
				}
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
				break;
			}
		}
		cbxPosition.SelectedItem = null;
	}

	private void A(XlOrientation A)
	{
		IEnumerator<Tuple<string, XlOrientation>> enumerator = default(IEnumerator<Tuple<string, XlOrientation>>);
		try
		{
			enumerator = LabelRotations.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Tuple<string, XlOrientation> current = enumerator.Current;
				if (current.Item2 != A)
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
					cbxRotation.SelectedItem = current;
					return;
				}
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_004d;
				}
				continue;
				end_IL_004d:
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
		cbxRotation.SelectedItem = null;
	}

	private void btnReset_Click(object sender, EventArgs e)
	{
		Series series = A();
		checked
		{
			if (B())
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
				Microsoft.Office.Interop.Excel.DataLabels dataLabels = (Microsoft.Office.Interop.Excel.DataLabels)series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
				XlDataLabelPosition position = dataLabels.Position;
				XlDataLabelPosition[] array = new XlDataLabelPosition[8]
				{
					XlDataLabelPosition.xlLabelPositionCenter,
					XlDataLabelPosition.xlLabelPositionAbove,
					XlDataLabelPosition.xlLabelPositionBelow,
					XlDataLabelPosition.xlLabelPositionLeft,
					XlDataLabelPosition.xlLabelPositionRight,
					XlDataLabelPosition.xlLabelPositionInsideBase,
					XlDataLabelPosition.xlLabelPositionInsideEnd,
					XlDataLabelPosition.xlLabelPositionOutsideEnd
				};
				int num = 0;
				while (true)
				{
					if (num < array.Length)
					{
						XlDataLabelPosition xlDataLabelPosition = array[num];
						if (position != xlDataLabelPosition)
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
								dataLabels.Position = xlDataLabelPosition;
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
								goto IL_008f;
							}
							break;
						}
						goto IL_008f;
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
					break;
					IL_008f:
					num++;
				}
				dataLabels.Position = position;
				XlOrientation xlOrientation = unchecked((XlOrientation)Conversions.ToInteger(dataLabels.Orientation));
				XlOrientation[] array2 = new XlOrientation[4]
				{
					XlOrientation.xlHorizontal,
					XlOrientation.xlDownward,
					XlOrientation.xlUpward,
					XlOrientation.xlVertical
				};
				int num2 = 0;
				while (true)
				{
					if (num2 < array2.Length)
					{
						XlOrientation xlOrientation2 = array2[num2];
						if (xlOrientation != xlOrientation2)
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
								dataLabels.Orientation = xlOrientation2;
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
								goto IL_0108;
							}
							break;
						}
						goto IL_0108;
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
					break;
					IL_0108:
					num2++;
				}
				dataLabels.Orientation = xlOrientation;
				dataLabels = null;
			}
			else
			{
				DataLabel dataLabel = (DataLabel)series.DataLabels(LabelDictionary[cbxLabel.SelectedIndex]);
				XlDataLabelPosition position = dataLabel.Position;
				XlDataLabelPosition[] array3 = new XlDataLabelPosition[8]
				{
					XlDataLabelPosition.xlLabelPositionCenter,
					XlDataLabelPosition.xlLabelPositionAbove,
					XlDataLabelPosition.xlLabelPositionBelow,
					XlDataLabelPosition.xlLabelPositionLeft,
					XlDataLabelPosition.xlLabelPositionRight,
					XlDataLabelPosition.xlLabelPositionInsideBase,
					XlDataLabelPosition.xlLabelPositionInsideEnd,
					XlDataLabelPosition.xlLabelPositionOutsideEnd
				};
				int num3 = 0;
				while (true)
				{
					if (num3 < array3.Length)
					{
						XlDataLabelPosition xlDataLabelPosition2 = array3[num3];
						if (position != xlDataLabelPosition2 || position == XlDataLabelPosition.xlLabelPositionCustom)
						{
							try
							{
								dataLabel.Position = xlDataLabelPosition2;
							}
							catch (Exception ex5)
							{
								ProjectData.SetProjectError(ex5);
								Exception ex6 = ex5;
								ProjectData.ClearProjectError();
								goto IL_01b0;
							}
							break;
						}
						goto IL_01b0;
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
					break;
					IL_01b0:
					num3++;
				}
				dataLabel.Position = position;
				XlOrientation xlOrientation = unchecked((XlOrientation)Conversions.ToInteger(dataLabel.Orientation));
				XlOrientation[] array4 = new XlOrientation[4]
				{
					XlOrientation.xlHorizontal,
					XlOrientation.xlDownward,
					XlOrientation.xlUpward,
					XlOrientation.xlVertical
				};
				int num4 = 0;
				while (true)
				{
					if (num4 < array4.Length)
					{
						XlOrientation xlOrientation3 = array4[num4];
						if (xlOrientation != xlOrientation3)
						{
							try
							{
								dataLabel.Orientation = xlOrientation3;
							}
							catch (Exception ex7)
							{
								ProjectData.SetProjectError(ex7);
								Exception ex8 = ex7;
								ProjectData.ClearProjectError();
								goto IL_0224;
							}
							break;
						}
						goto IL_0224;
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
					break;
					IL_0224:
					num4++;
				}
				dataLabel.Orientation = xlOrientation;
				dataLabel = null;
			}
			series = null;
		}
	}

	private void wpfMoveDataLabels_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (txtStep.IsKeyboardFocused)
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
			if (cbxLabel.IsKeyboardFocused || cbxPosition.IsKeyboardFocused)
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
				if (cbxRotation.IsKeyboardFocused || cbxSeries.IsKeyboardFocused)
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
					switch (e.Key)
					{
					case Key.Up:
						A(A);
						e.Handled = true;
						break;
					case Key.Down:
						A(B);
						e.Handled = true;
						break;
					case Key.Left:
						A(C);
						e.Handled = true;
						break;
					case Key.Right:
						A(D);
						e.Handled = true;
						break;
					case Key.D1:
					case Key.NumPad1:
						txtStep.Text = Conversions.ToString(1);
						e.Handled = true;
						break;
					case Key.D2:
					case Key.NumPad2:
						txtStep.Text = Conversions.ToString(2);
						e.Handled = true;
						break;
					case Key.D3:
					case Key.NumPad3:
						txtStep.Text = Conversions.ToString(3);
						e.Handled = true;
						break;
					case Key.D4:
					case Key.NumPad4:
						txtStep.Text = Conversions.ToString(4);
						e.Handled = true;
						break;
					case Key.D5:
					case Key.NumPad5:
						txtStep.Text = Conversions.ToString(5);
						e.Handled = true;
						break;
					case Key.D6:
					case Key.NumPad6:
						txtStep.Text = Conversions.ToString(6);
						e.Handled = true;
						break;
					case Key.D7:
					case Key.NumPad7:
						txtStep.Text = Conversions.ToString(7);
						e.Handled = true;
						break;
					case Key.D8:
					case Key.NumPad8:
						txtStep.Text = Conversions.ToString(8);
						e.Handled = true;
						break;
					case Key.D9:
					case Key.NumPad9:
						txtStep.Text = Conversions.ToString(9);
						e.Handled = true;
						break;
					case Key.F2:
						txtStep.Focus();
						txtStep.SelectAll();
						e.Handled = true;
						break;
					}
					return;
				}
			}
		}
	}

	private void ArrowMouseDown(object sender, MouseButtonEventArgs e)
	{
		A((System.Windows.Controls.Button)sender);
	}

	private void ArrowMouseUp(object sender, MouseButtonEventArgs e)
	{
		G();
	}

	private void ArrowKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Space)
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
			if (e.IsRepeat)
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
				A((System.Windows.Controls.Button)sender);
				return;
			}
		}
	}

	private void ArrowKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Space)
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
			G();
			return;
		}
	}

	private void A(System.Windows.Controls.Button A)
	{
		this.m_C = 0;
		string name = A.Name;
		if (Operators.CompareString(name, btnUp.Name, TextCompare: false) == 0)
		{
			MoveAction = this.A;
		}
		else if (Operators.CompareString(name, btnDown.Name, TextCompare: false) == 0)
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
			MoveAction = B;
		}
		else if (Operators.CompareString(name, btnLeft.Name, TextCompare: false) == 0)
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
			MoveAction = C;
		}
		else
		{
			MoveAction = D;
		}
		timer.Start();
	}

	private void G()
	{
		timer.Stop();
		if (this.m_C > 3)
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
			A(MoveAction);
			return;
		}
	}

	private void MoveRepeat(object sender, EventArgs e)
	{
		checked
		{
			this.m_C++;
			if (this.m_C <= 3)
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
				A(MoveAction);
				return;
			}
		}
	}

	private void A(Action<DataLabel, int> A)
	{
		if (int.TryParse(txtStep.Text, out var result))
		{
			if (result >= 1)
			{
				Series series = this.A();
				if (B())
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
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = ((IEnumerable)series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						while (enumerator.MoveNext())
						{
							DataLabel arg = (DataLabel)enumerator.Current;
							try
							{
								A(arg, result);
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
				}
				else
				{
					try
					{
						A((DataLabel)series.DataLabels(LabelDictionary[cbxLabel.SelectedIndex]), result);
						cbxPosition.SelectedIndex = -1;
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
				}
				series = null;
				return;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
		}
		Forms.WarningMessage(VH.A(57011));
		txtStep.Focus();
		txtStep.SelectAll();
	}

	private void A(DataLabel A, int B)
	{
		A.Top -= B;
		this.B(A, 0, checked(-B));
	}

	private void B(DataLabel A, int B)
	{
		A.Top += B;
		this.B(A, 0, B);
	}

	private void C(DataLabel A, int B)
	{
		A.Left -= B;
		this.B(A, checked(-B), 0);
	}

	private void D(DataLabel A, int B)
	{
		A.Left += B;
		this.B(A, B, 0);
	}

	private void B(DataLabel A, int B, int C)
	{
		Dictionary<DataLabel, System.Windows.Point> movedLabels = MovedLabels;
		if (movedLabels.TryGetValue(A, out var value))
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
			movedLabels[A] = new System.Windows.Point(value.X + (double)B, value.Y + (double)C);
		}
		else
		{
			movedLabels.Add(A, new System.Windows.Point(B, C));
		}
		movedLabels = null;
	}

	private void H()
	{
		cbxPosition.SelectionChanged += PositionSelectionChanged;
	}

	private void I()
	{
		cbxPosition.SelectionChanged -= PositionSelectionChanged;
	}

	private void J()
	{
		cbxRotation.SelectionChanged += RotationSelectionChanged;
	}

	private void K()
	{
		cbxRotation.SelectionChanged -= RotationSelectionChanged;
	}

	private void L()
	{
		H();
		J();
	}

	private void M()
	{
		I();
		K();
	}

	private void N()
	{
		cbxSeries.SelectionChanged -= SeriesSelectionChanged;
		cbxLabel.SelectionChanged -= LabelSelectionChanged;
		M();
	}

	private bool A(Series A)
	{
		bool result;
		try
		{
			_ = (Microsoft.Office.Interop.Excel.DataLabels)A.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
			goto IL_00cd;
		}
		checked
		{
			try
			{
				Points points = (Points)A.Points(RuntimeHelpers.GetObjectValue(Missing.Value));
				int count = points.Count;
				int num = count;
				while (true)
				{
					if (num < 1)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							points = null;
							break;
						}
						break;
					}
					if (points.Item(num).HasDataLabel)
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
							points = null;
							result = true;
							break;
						}
					}
					else
					{
						if (num >= count - 100)
						{
							num += -1;
							continue;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							points = null;
							result = A.HasDataLabels;
							break;
						}
					}
					goto end_IL_0030;
				}
				goto IL_00cb;
				end_IL_0030:;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
				goto IL_00cb;
			}
			goto IL_00cd;
		}
		IL_00cd:
		return result;
		IL_00cb:
		result = false;
		goto IL_00cd;
	}

	private void C(string A)
	{
		Forms.ErrorMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void D(string A)
	{
		Forms.WarningMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void E(string A)
	{
		Forms.InfoMessage(System.Windows.Window.GetWindow(this), A);
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (this.m_D)
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
			this.m_D = true;
			Uri resourceLocator = new Uri(VH.A(57114), UriKind.Relative);
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
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
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
					scroller = (ScrollViewer)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					stkContents = (StackPanel)target;
					return;
				}
			}
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
					cbxSeries = (System.Windows.Controls.ComboBox)target;
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
					cbxLabel = (System.Windows.Controls.ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			cbxPosition = (System.Windows.Controls.ComboBox)target;
			return;
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
					btnUp = (System.Windows.Controls.Button)target;
					return;
				}
			}
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
					btnLeft = (System.Windows.Controls.Button)target;
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
					txtStep = (System.Windows.Controls.TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			btnRight = (System.Windows.Controls.Button)target;
			return;
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
					btnDown = (System.Windows.Controls.Button)target;
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
					cbxRotation = (System.Windows.Controls.ComboBox)target;
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
			btnClose = (System.Windows.Controls.Button)target;
			break;
		case 14:
			btnDismiss = (System.Windows.Controls.Button)target;
			break;
		default:
			this.m_D = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
