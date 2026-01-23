using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Xml;
using A;
using ExcelAddIn1.Formulas;
using Foo.Controls;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts.GrowthArrow;

[DesignerGenerated]
public sealed class wpfPane : System.Windows.Controls.UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private ArrowOptions m_A;

	private Microsoft.Office.Interop.Excel.Application m_A;

	private bool m_A;

	private readonly int m_A;

	[CompilerGenerated]
	private Chart m_A;

	[CompilerGenerated]
	private List<double> m_A;

	private ObservableCollection<Arrow> m_A;

	private Dictionary<int, string> m_A;

	private string m_A;

	private Visibility m_A;

	private double m_A;

	private double m_B;

	private double m_C;

	private double m_D;

	private int m_B;

	[CompilerGenerated]
	private bool m_B;

	[AccessedThroughProperty("scroller")]
	[CompilerGenerated]
	private ScrollViewer m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("stkContents")]
	private StackPanel m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("grpArrows")]
	private System.Windows.Controls.GroupBox m_A;

	[AccessedThroughProperty("btnAddArrow")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnRecommend")]
	private System.Windows.Controls.Button m_B;

	[AccessedThroughProperty("numOffset")]
	[CompilerGenerated]
	private MacNumericUpDown m_A;

	[AccessedThroughProperty("numHeight")]
	[CompilerGenerated]
	private MacNumericUpDown m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnColor")]
	private System.Windows.Controls.Button m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("numWeight")]
	private MacNumericUpDown m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxStyle")]
	private System.Windows.Controls.ComboBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxPosition")]
	private System.Windows.Controls.ComboBox m_B;

	[AccessedThroughProperty("cbxShape")]
	[CompilerGenerated]
	private System.Windows.Controls.ComboBox m_C;

	[AccessedThroughProperty("chkBorder")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_A;

	[AccessedThroughProperty("chkFill")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_B;

	[AccessedThroughProperty("chkRotate")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("chkBold")]
	private System.Windows.Controls.CheckBox m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("chkMatch")]
	private System.Windows.Controls.CheckBox m_E;

	[AccessedThroughProperty("txtFormat")]
	[CompilerGenerated]
	private System.Windows.Controls.TextBox m_A;

	[AccessedThroughProperty("btnClose")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_D;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private System.Windows.Controls.Button m_F;

	[AccessedThroughProperty("btnDismiss")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_G;

	private bool m_C;

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

	private List<double> ChartTotals
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

	public ObservableCollection<Arrow> GrowthArrows
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(53024));
		}
	}

	public Dictionary<int, string> Periods
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(53049));
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
			ErrorVisibility = ((value.Length <= 0) ? Visibility.Collapsed : Visibility.Visible);
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

	public double YOffset
	{
		get
		{
			return this.m_A;
		}
		set
		{
			if (!IsLoading)
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
				if (value != this.m_A)
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
					O(this.m_A, value);
				}
			}
			this.m_A = value;
			A(VH.A(53114));
		}
	}

	public double YRise
	{
		get
		{
			return this.m_B;
		}
		set
		{
			if (!IsLoading)
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
				if (value != this.m_B)
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
					P(this.m_B, value);
				}
			}
			this.m_B = value;
			A(VH.A(53129));
		}
	}

	public double LargeChange
	{
		get
		{
			return this.m_C;
		}
		set
		{
			this.m_C = value;
			A(VH.A(53140));
		}
	}

	public double SmallChange
	{
		get
		{
			return this.m_D;
		}
		set
		{
			this.m_D = value;
			A(VH.A(53163));
		}
	}

	public int NumberDecimalDigits
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(VH.A(53186));
		}
	}

	private bool IsLoading
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

	internal virtual System.Windows.Controls.GroupBox grpArrows
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

	internal virtual System.Windows.Controls.Button btnAddArrow
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
			RoutedEventHandler value2 = AddArrow;
			System.Windows.Controls.Button button = this.m_A;
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

	internal virtual System.Windows.Controls.Button btnRecommend
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
			RoutedEventHandler value2 = RecommendArrows;
			System.Windows.Controls.Button button = this.m_B;
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
			this.m_B = value;
			button = this.m_B;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual MacNumericUpDown numOffset
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

	internal virtual MacNumericUpDown numHeight
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

	internal virtual System.Windows.Controls.Button btnColor
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
			RoutedEventHandler value2 = LineColorChanged;
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
				switch (5)
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

	internal virtual MacNumericUpDown numWeight
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

	internal virtual System.Windows.Controls.ComboBox cbxStyle
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

	internal virtual System.Windows.Controls.ComboBox cbxPosition
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

	internal virtual System.Windows.Controls.ComboBox cbxShape
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

	internal virtual System.Windows.Controls.CheckBox chkBorder
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

	internal virtual System.Windows.Controls.CheckBox chkFill
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

	internal virtual System.Windows.Controls.CheckBox chkRotate
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

	internal virtual System.Windows.Controls.CheckBox chkBold
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

	internal virtual System.Windows.Controls.CheckBox chkMatch
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

	internal virtual System.Windows.Controls.TextBox txtFormat
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

	internal virtual System.Windows.Controls.Button btnClose
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
			RoutedEventHandler value2 = btnClose_Click;
			System.Windows.Controls.Button button = this.m_D;
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
			RoutedEventHandler value2 = btnCancel_Click;
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

	internal virtual System.Windows.Controls.Button btnOk
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
			RoutedEventHandler value2 = btnOk_Click;
			System.Windows.Controls.Button button = this.m_F;
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
			this.m_F = value;
			button = this.m_F;
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
			this.m_G = value;
			button = this.m_G;
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
		this.m_A = ColorTranslator.ToOle(System.Drawing.Color.White);
		this.m_A = "";
		this.m_A = Visibility.Collapsed;
		this.m_A = 0.0;
		this.m_B = 0.0;
		this.m_C = 1.0;
		this.m_D = 1.0;
		this.m_B = 0;
		IsLoading = false;
		InitializeComponent();
		Dictionary<string, Geometry> itemsSource = new Dictionary<string, Geometry>
		{
			{
				VH.A(55655),
				Geometry.Parse(VH.A(55666))
			},
			{
				VH.A(55932),
				Geometry.Parse(VH.A(55945))
			},
			{
				VH.A(56075),
				Geometry.Parse(VH.A(56084))
			}
		};
		cbxStyle.ItemsSource = itemsSource;
		itemsSource = null;
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

	private void wpfPane_Unloaded(object sender, RoutedEventArgs e)
	{
	}

	private void PaneSizeChanged(object sender, SizeChangedEventArgs e)
	{
		Panes.PaneSizeChanged(scroller, e);
	}

	private void ContentsHeightChanged(object sender, SizeChangedEventArgs e)
	{
		if (!IsLoading && e.HeightChanged)
		{
			Panes.UpdateScrollViewerPadding(scroller);
		}
	}

	public void ShowPane()
	{
		this.m_A = MH.A.Application;
		this.m_A = A();
		this.m_A = false;
		GrowthArrows = new ObservableCollection<Arrow>();
		Periods = new Dictionary<int, string>();
		base.SizeChanged -= PaneSizeChanged;
		base.SizeChanged += PaneSizeChanged;
		stkContents.SizeChanged -= ContentsHeightChanged;
		stkContents.SizeChanged += ContentsHeightChanged;
		B();
	}

	public void HidePane()
	{
		M();
		base.SizeChanged -= PaneSizeChanged;
		stkContents.SizeChanged -= ContentsHeightChanged;
		if (!this.m_A)
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
			C();
		}
		GrowthArrows = null;
		Periods = null;
		Chart = null;
		ChartTotals = null;
		this.m_A = null;
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		N();
	}

	private void btnCancel_Click(object sender, RoutedEventArgs e)
	{
		B();
	}

	private void btnClose_Click(object sender, RoutedEventArgs e)
	{
		A();
	}

	private void DismissError(object sender, RoutedEventArgs e)
	{
		D();
	}

	private void HandleKeyPresses(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key == Key.Return)
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
					if (A())
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								A();
								return;
							}
						}
					}
					N();
					return;
				}
			}
		}
		if (e.Key == Key.Escape)
		{
			C();
		}
	}

	private void A()
	{
		X();
		D();
	}

	private void B()
	{
		E();
		D();
	}

	private void C()
	{
		if (A())
		{
			A();
		}
		else
		{
			B();
		}
	}

	private void D()
	{
		this.m_A = true;
		Pane.B();
		Pane.C();
	}

	private void E()
	{
		this.m_A.ScreenUpdating = false;
		F();
		this.m_A.ScreenUpdating = true;
	}

	private void F()
	{
		IEnumerator<Arrow> enumerator = default(IEnumerator<Arrow>);
		try
		{
			enumerator = GrowthArrows.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Arrow current = enumerator.Current;
				A(current.Series);
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
		GrowthArrows.Clear();
	}

	private void A(Series A)
	{
		try
		{
			A.Delete();
			Thread.Sleep(500);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private bool A()
	{
		return btnClose.Visibility == Visibility.Visible;
	}

	private void G()
	{
		try
		{
			new ComAwareEventInfo(typeof(ChartEvents_Event), VH.A(53225)).RemoveEventHandler(Chart, new ChartEvents_DeactivateEventHandler(G));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (A())
		{
			if (!B())
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
						A();
						return;
					}
				}
			}
			H();
			return;
		}
		E();
		if (!B())
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					D();
					return;
				}
			}
		}
		H();
	}

	private void H()
	{
		this.m_A.CommandBars.ReleaseFocus();
	}

	private void B(string A)
	{
		ErrorText = A;
	}

	private void I()
	{
		ErrorText = "";
	}

	private bool B()
	{
		bool B = false;
		bool flag = false;
		ChartTotals = null;
		Chart = Helpers.SelectedChart();
		checked
		{
			if (Chart != null)
			{
				{
					IEnumerator enumerator = ((IEnumerable)Chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
					try
					{
						while (true)
						{
							if (enumerator.MoveNext())
							{
								Series series = (Series)enumerator.Current;
								Series series2 = series;
								if (A(series, ref B))
								{
									int num = Conversions.ToInteger(NewLateBinding.LateGet(series2.Values, null, VH.A(53246), new object[0], null, null, null));
									if (ChartTotals == null)
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
										ChartTotals = new List<double>();
										int num2 = num;
										for (int i = 1; i <= num2; i++)
										{
											ChartTotals.Add(0.0);
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
									int num3 = num;
									for (int j = 1; j <= num3; j++)
									{
										List<double> chartTotals;
										int index;
										(chartTotals = ChartTotals)[index = j - 1] = Conversions.ToDouble(Operators.AddObject(chartTotals[index], NewLateBinding.LateIndexGet(series2.Values, new object[1] { j }, null)));
									}
								}
								series2 = null;
								if (!A(series))
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
								if (this.B(series))
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
									GrowthArrows.Add(new Arrow(series));
									continue;
								}
								this.B(VH.A(53259));
								flag = true;
								break;
							}
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
									goto end_IL_01c1;
								}
								continue;
								end_IL_01c1:
								break;
							}
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
				new ComAwareEventInfo(typeof(ChartEvents_Event), VH.A(53225)).RemoveEventHandler(Chart, new ChartEvents_DeactivateEventHandler(G));
				new ComAwareEventInfo(typeof(ChartEvents_Event), VH.A(53225)).AddEventHandler(Chart, new ChartEvents_DeactivateEventHandler(G));
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
					if (ChartTotals != null)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							I();
							J();
							clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(52976));
							return true;
						}
					}
					Chart = null;
					this.B(VH.A(53565));
				}
			}
			else
			{
				this.B(VH.A(53749));
			}
			return false;
		}
	}

	private bool A(Series A)
	{
		int num;
		if (A.ChartType == XlChartType.xlXYScatterLinesNoMarkers)
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
			if (A.AxisGroup == XlAxisGroup.xlPrimary)
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
				num = ((Conversions.ToBoolean(Operators.CompareObjectEqual(NewLateBinding.LateGet(A.Values, null, VH.A(53246), new object[0], null, null, null), 5, TextCompare: false)) || Conversions.ToBoolean(Operators.CompareObjectEqual(NewLateBinding.LateGet(A.Values, null, VH.A(53246), new object[0], null, null, null), 3, TextCompare: false))) ? 1 : 0);
				goto IL_00b2;
			}
		}
		num = 0;
		goto IL_00b2;
		IL_00b2:
		return Conversions.ToBoolean((byte)num != 0);
	}

	private bool A(Series A, ref bool B)
	{
		Series series = A;
		if (series.ChartType == XlChartType.xlColumnStacked)
		{
			goto IL_0041;
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
		if (series.ChartType == XlChartType.xlColumnClustered)
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
			if (!B)
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
				goto IL_0041;
			}
		}
		goto IL_006f;
		IL_006f:
		int num = 0;
		goto IL_0096;
		IL_0041:
		if (series.AxisGroup != XlAxisGroup.xlPrimary)
		{
			goto IL_006f;
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
		num = ((series.Format.Fill.Visible == MsoTriState.msoTrue) ? 1 : 0);
		if (num != 0)
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
			if (series.ChartType == XlChartType.xlColumnClustered)
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
				B = true;
			}
		}
		goto IL_0096;
		IL_0096:
		series = null;
		return (byte)num != 0;
	}

	private void J()
	{
		Axis axis = (Axis)Chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue);
		IsLoading = true;
		M();
		int count = ChartTotals.Count;
		btnAddArrow.IsEnabled = count > 2;
		System.Windows.Controls.Button button = btnRecommend;
		int visibility;
		if (count <= 2)
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
			visibility = 2;
		}
		else
		{
			visibility = 0;
		}
		button.Visibility = (Visibility)visibility;
		object objectValue = RuntimeHelpers.GetObjectValue(((Axis)Chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory)).CategoryNames);
		int num = count;
		for (int i = 1; i <= num; i = checked(i + 1))
		{
			try
			{
				Periods.Add(i, NewLateBinding.LateIndexGet(objectValue, new object[1] { i }, null).ToString());
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Periods.Add(i, VH.A(53848) + i);
				ProjectData.ClearProjectError();
			}
		}
		IEnumerator<Arrow> enumerator = default(IEnumerator<Arrow>);
		IEnumerator<Arrow> enumerator2 = default(IEnumerator<Arrow>);
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			try
			{
				bool flag;
				if (GrowthArrows.Count == 0)
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
					YOffset = axis.MajorUnit;
					YRise = YOffset;
					A();
					U();
					flag = false;
				}
				else
				{
					try
					{
						enumerator = GrowthArrows.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Arrow current = enumerator.Current;
							object objectValue2 = RuntimeHelpers.GetObjectValue(current.Series.Values);
							if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(objectValue2, null, VH.A(53246), new object[0], null, null, null), 5, TextCompare: false))
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
								this.m_A.LineType = ArrowType.Elbow;
								YOffset = Conversions.ToDouble(Operators.SubtractObject(NewLateBinding.LateIndexGet(objectValue2, new object[1] { 1 }, null), ChartTotals[0]));
								YRise = Conversions.ToDouble(NewLateBinding.LateGet(null, typeof(Math), VH.A(53859), new object[2]
								{
									Operators.SubtractObject(NewLateBinding.LateIndexGet(objectValue2, new object[1] { 2 }, null), NewLateBinding.LateIndexGet(objectValue2, new object[1] { 1 }, null)),
									Operators.SubtractObject(NewLateBinding.LateIndexGet(objectValue2, new object[1] { 4 }, null), NewLateBinding.LateIndexGet(objectValue2, new object[1] { 5 }, null))
								}, null, null, null));
								current.A(3);
							}
							else
							{
								if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateIndexGet(objectValue2, new object[1] { 1 }, null), NewLateBinding.LateIndexGet(objectValue2, new object[1] { NewLateBinding.LateGet(objectValue2, null, VH.A(53246), new object[0], null, null, null) }, null), TextCompare: false))
								{
									this.m_A.LineType = ArrowType.Flat;
									YOffset = Conversions.ToDouble(Operators.SubtractObject(NewLateBinding.LateIndexGet(objectValue2, new object[1] { 1 }, null), ChartTotals.Max()));
									YRise = 0.0;
								}
								else
								{
									this.m_A.LineType = ArrowType.Angled;
									YOffset = Conversions.ToDouble(Operators.SubtractObject(NewLateBinding.LateIndexGet(objectValue2, new object[1] { 1 }, null), ChartTotals[0]));
									YRise = 0.0;
								}
								current.A(2);
							}
							B(current);
							if (current.Label == null)
							{
								this.m_A.LabelPosition = CagrLabelPosition.None;
								Q();
							}
							else
							{
								A(current);
							}
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_0425;
							}
							continue;
							end_IL_0425:
							break;
						}
					}
					finally
					{
						if (enumerator != null)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								enumerator.Dispose();
								break;
							}
						}
					}
					V();
					flag = true;
				}
				K();
				LargeChange = axis.MajorUnit;
				SmallChange = axis.MinorUnit;
				NumberDecimalDigits = ((axis.MinorUnit % 1.0 != 0.0) ? 1 : 0);
				T();
				if (flag)
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
					if (chkRotate.IsEnabled)
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
						if (this.m_A.Rotate)
						{
							try
							{
								enumerator2 = GrowthArrows.GetEnumerator();
								while (enumerator2.MoveNext())
								{
									Arrow current2 = enumerator2.Current;
									if (current2.Label != null)
									{
										current2.A();
									}
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
										goto end_IL_050f;
									}
									continue;
									end_IL_050f:
									break;
								}
							}
							finally
							{
								if (enumerator2 != null)
								{
									while (true)
									{
										switch (1)
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
					}
				}
				if (!flag)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							btnOk.Focus();
							return;
						}
					}
				}
				btnClose.Focus();
				return;
			}
			finally
			{
				axis = null;
				IsLoading = false;
				L();
			}
		}
	}

	private bool B(Series A)
	{
		return A.Formula.Contains(VH.A(41262));
	}

	private void A(Arrow A)
	{
		int rGB = A.Series.Format.Line.ForeColor.RGB;
		ChartFormat format = A.Label.Format;
		Font2 font = format.TextFrame2.TextRange.Font;
		ref ArrowOptions a = ref this.m_A;
		a.LabelBold = font.Bold == MsoTriState.msoTrue;
		a.LabelBorder = format.Line.Visible == MsoTriState.msoTrue;
		a.Invert = format.Fill.ForeColor.RGB == rGB;
		txtFormat.Text = "";
		if (font.Fill.Visible == MsoTriState.msoFalse)
		{
			a.LabelPosition = CagrLabelPosition.None;
			Q();
		}
		else
		{
			XlDataLabelPosition position = A.Label.Position;
			if (position != XlDataLabelPosition.xlLabelPositionCenter)
			{
				if (position != XlDataLabelPosition.xlLabelPositionAbove)
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
					if (position != XlDataLabelPosition.xlLabelPositionBelow)
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
						a.LabelPosition = this.m_A.LabelPosition;
					}
					else
					{
						a.LabelPosition = CagrLabelPosition.BelowLine;
					}
				}
				else
				{
					a.LabelPosition = CagrLabelPosition.AboveLine;
				}
			}
			else
			{
				a.LabelPosition = CagrLabelPosition.Inline;
			}
		}
		a.MatchColor = font.Fill.ForeColor.RGB == rGB;
		a.Rotate = Conversions.ToBoolean(Conversions.ToBoolean(Operators.CompareObjectNotEqual(A.Label.Orientation, XlOrientation.xlHorizontal, TextCompare: false)) && Conversions.ToBoolean(Operators.CompareObjectNotEqual(A.Label.Orientation, 0, TextCompare: false)));
		a.Shape = format.AutoShapeType;
		txtFormat.Text = this.m_A.Format;
		font = null;
		format = null;
	}

	private void K()
	{
		ref ArrowOptions a = ref this.m_A;
		chkBold.IsChecked = a.LabelBold;
		chkBorder.IsChecked = a.LabelBorder;
		chkFill.IsChecked = a.Invert;
		System.Drawing.Color color = ColorTranslator.FromOle(a.Color);
		btnColor.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(color.R, color.G, color.B));
		color = default(System.Drawing.Color);
		cbxPosition.SelectedIndex = (int)a.LabelPosition;
		if (cbxPosition.SelectedIndex == 1)
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
			cbxShape.IsEnabled = false;
			chkBorder.IsEnabled = false;
			chkFill.IsEnabled = false;
		}
		else if (a.LabelPosition != CagrLabelPosition.None)
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
			W();
		}
		chkMatch.IsChecked = a.MatchColor;
		chkRotate.IsChecked = a.Rotate;
		numWeight.Value = a.Weight;
	}

	private void L()
	{
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_0030: Expected O, but got Unknown
		cbxStyle.SelectionChanged += LineStyleChanged;
		numWeight.ValueChanged += new MacRangeBaseValueChangedHandler(LineWeightChanged);
		cbxPosition.SelectionChanged += LabelPositionChanged;
		cbxShape.SelectionChanged += ShapeTypeChanged;
		txtFormat.TextChanged += txtFormat_TextChanged;
		txtFormat.LostFocus += txtFormat_LostFocus;
		chkBold.Checked += chkBold_CheckedChanged;
		chkBold.Unchecked += chkBold_CheckedChanged;
		chkBorder.Checked += chkBorder_CheckedChanged;
		chkBorder.Unchecked += chkBorder_CheckedChanged;
		chkFill.Checked += chkFill_CheckedChanged;
		chkFill.Unchecked += chkFill_CheckedChanged;
		chkRotate.Checked += chkRotate_CheckedChanged;
		chkRotate.Unchecked += chkRotate_CheckedChanged;
		chkMatch.Checked += chkMatch_CheckedChanged;
		chkMatch.Unchecked += chkMatch_CheckedChanged;
	}

	private void M()
	{
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_0030: Expected O, but got Unknown
		cbxStyle.SelectionChanged -= LineStyleChanged;
		numWeight.ValueChanged -= new MacRangeBaseValueChangedHandler(LineWeightChanged);
		cbxPosition.SelectionChanged -= LabelPositionChanged;
		cbxShape.SelectionChanged -= ShapeTypeChanged;
		txtFormat.TextChanged -= txtFormat_TextChanged;
		txtFormat.LostFocus -= txtFormat_LostFocus;
		chkBold.Checked -= chkBold_CheckedChanged;
		chkBold.Unchecked -= chkBold_CheckedChanged;
		chkBorder.Checked -= chkBorder_CheckedChanged;
		chkBorder.Unchecked -= chkBorder_CheckedChanged;
		chkFill.Checked -= chkFill_CheckedChanged;
		chkFill.Unchecked -= chkFill_CheckedChanged;
		chkRotate.Checked -= chkRotate_CheckedChanged;
		chkRotate.Unchecked -= chkRotate_CheckedChanged;
		chkMatch.Checked -= chkMatch_CheckedChanged;
		chkMatch.Unchecked -= chkMatch_CheckedChanged;
	}

	private Arrow A()
	{
		return B(1, ChartTotals.Count);
	}

	private Arrow B(int A, int B)
	{
		Arrow arrow = null;
		this.m_A.ScreenUpdating = false;
		Axis axis;
		try
		{
			axis = (Axis)Chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue);
			if (((UIElement)(object)numHeight).IsEnabled)
			{
				IsLoading = true;
				YRise = axis.MajorUnit;
				IsLoading = false;
			}
			arrow = new Arrow(((SeriesCollection)Chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).NewSeries(), A, B);
			arrow.A(ChartTotals, this.A());
			GrowthArrows.Add(arrow);
			Series series = arrow.Series;
			series.ChartType = XlChartType.xlXYScatterLinesNoMarkers;
			series.AxisGroup = XlAxisGroup.xlPrimary;
			series.Name = VH.A(53866);
			arrow.A(ChartTotals, YOffset, YRise, this.m_A);
			LineFormat line = series.Format.Line;
			line.ForeColor.RGB = this.m_A.Color;
			line.Weight = this.m_A.Weight;
			line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
			line.EndArrowheadLength = MsoArrowheadLength.msoArrowheadLong;
			line.EndArrowheadWidth = MsoArrowheadWidth.msoArrowheadWidthMedium;
			_ = null;
			txtFormat.Text = this.m_A.Format;
			C(arrow);
			if (this.m_A.Rotate)
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
				arrow.A();
			}
			if (this.m_A.LabelPosition == CagrLabelPosition.None)
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
				this.A(arrow, B: false);
				Q();
			}
			_ = null;
			if (!axis.MaximumScaleIsAuto)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					Forms.InfoMessage(VH.A(53879));
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			C(ex2.Message);
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		this.m_A.ScreenUpdating = true;
		axis = null;
		return arrow;
	}

	private void A(DataLabel A)
	{
		ChartFormat format = A.Format;
		if (this.m_A.Shape == (MsoAutoShapeType)0)
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
			if (this.m_A.LabelPosition == CagrLabelPosition.Inline)
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
				W();
			}
		}
		if (this.m_A.LabelPosition == CagrLabelPosition.Inline)
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
			if (this.m_A.Shape == (MsoAutoShapeType)0)
			{
				this.m_A.Shape = MsoAutoShapeType.msoShapeRectangle;
			}
			format.AutoShapeType = this.m_A.Shape;
		}
		Font2 font = format.TextFrame2.TextRange.Font;
		if (this.m_A.LabelBold)
		{
			font.Bold = MsoTriState.msoTrue;
		}
		if (this.m_A.Invert)
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
			font.Fill.ForeColor.RGB = this.m_A;
		}
		else if (this.m_A.MatchColor)
		{
			font.Fill.ForeColor.RGB = this.m_A.Color;
		}
		font = null;
		if (this.m_A.LabelPosition == CagrLabelPosition.Inline)
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
			if (this.m_A.LabelBorder)
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
				LineFormat line = format.Line;
				line.Weight = this.m_A.Weight;
				line.Visible = MsoTriState.msoTrue;
				line.ForeColor.RGB = this.m_A.Color;
				_ = null;
			}
			format.Fill.Visible = MsoTriState.msoTrue;
			if (this.m_A.Invert)
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
				format.Fill.ForeColor.RGB = this.m_A.Color;
			}
			else
			{
				format.Fill.ForeColor.RGB = this.m_A;
			}
		}
		format = null;
	}

	private void B(Arrow A)
	{
		ref ArrowOptions a = ref this.m_A;
		a.Color = A.Series.Format.Line.ForeColor.RGB;
		a.Format = this.m_A.Format;
		a.Weight = A.Series.Format.Line.Weight;
	}

	private void C(Arrow A)
	{
		int num = ((this.m_A.LineType == ArrowType.Elbow) ? 3 : 2);
		A.A(this.m_A, num);
		this.A(((Microsoft.Office.Interop.Excel.Point)A.Series.Points(num)).DataLabel);
	}

	private void N()
	{
		ChartObject chartObject = (ChartObject)Chart.Parent;
		int c = QuickCharts2.InputColor();
		bool d = cbxStyle.SelectedIndex == 0;
		string name = chartObject.TopLeftCell.Worksheet.Name;
		Range B = null;
		List<Range> C = null;
		string D = default(string);
		try
		{
			A(name, ref B, ref C, ref D);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			chartObject = null;
			ProjectData.ClearProjectError();
			return;
		}
		Microsoft.Office.Interop.Excel.Application a = this.m_A;
		XlCalculation calculation = a.Calculation;
		bool autoExpandListRange = a.AutoCorrect.AutoExpandListRange;
		a.AutoCorrect.AutoExpandListRange = false;
		a.ScreenUpdating = false;
		a.EnableEvents = false;
		a.Calculation = XlCalculation.xlCalculationManual;
		_ = null;
		XlPlacement xlPlacement = (XlPlacement)Conversions.ToInteger(chartObject.Placement);
		chartObject.Placement = XlPlacement.xlMove;
		try
		{
			IEnumerator<Arrow> enumerator = default(IEnumerator<Arrow>);
			try
			{
				enumerator = GrowthArrows.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Arrow current = enumerator.Current;
					A(chartObject, current, c, d, B, C, name, D);
					if (this.m_A.LabelPosition != CagrLabelPosition.None)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					try
					{
						current.Series.HasDataLabels = false;
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_014e;
					}
					continue;
					end_IL_014e:
					break;
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
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			this.C(ex6.Message);
			clsReporting.LogException(ex6);
			ProjectData.ClearProjectError();
		}
		try
		{
			if (GrowthArrows.Count == 1)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					GrowthArrows[0].Series.Select();
					break;
				}
			}
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
		chartObject.Placement = xlPlacement;
		chartObject = null;
		Microsoft.Office.Interop.Excel.Application a2 = this.m_A;
		a2.AutoCorrect.AutoExpandListRange = autoExpandListRange;
		if (calculation == XlCalculation.xlCalculationManual)
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
			a2.Calculation = XlCalculation.xlCalculationSemiautomatic;
		}
		else
		{
			a2.Calculation = calculation;
		}
		a2.ScreenUpdating = true;
		a2.EnableEvents = true;
		a2 = null;
		B = null;
		C = null;
		X();
		V();
		btnClose.Focus();
	}

	private void A(string A, ref Range B, ref List<Range> C, ref string D)
	{
		string[] array = null;
		bool B2 = false;
		Range range = null;
		checked
		{
			try
			{
				array = Helpers.A(((Series)Chart.SeriesCollection(1)).Formula);
				if (!modFunctionsStr.IsBlank(array[1]))
				{
					B = ((_Application)this.m_A).get_Range((object)array[1], RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				C = new List<Range>();
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = ((IEnumerable)Chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
					while (enumerator.MoveNext())
					{
						Series series = (Series)enumerator.Current;
						Series series2 = series;
						if (this.A(series, ref B2))
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
							array = Helpers.A(series2.Formula);
							range = ((_Application)this.m_A).get_Range((object)array[2], RuntimeHelpers.GetObjectValue(Missing.Value));
							int num = Conversions.ToInteger(range.Cells.CountLarge);
							for (int i = 1; i <= num; i++)
							{
								if (Operators.ConditionalCompareObjectGreaterEqual(C.Count, range.Cells.CountLarge, TextCompare: false))
								{
									C[i - 1] = this.m_A.Union(C[i - 1], (Range)range.Cells[i, RuntimeHelpers.GetObjectValue(Missing.Value)], RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								}
								else
								{
									C.Add((Range)range.Cells[i, RuntimeHelpers.GetObjectValue(Missing.Value)]);
								}
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
						series2 = null;
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
				string text = array[2].Split('$')[0];
				if (range == null || Operators.CompareString(A + VH.A(7827), text, TextCompare: false) == 0)
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
					D = text;
					return;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				if (array != null)
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
					if (array[2].Contains(Conversions.ToString(Path.DirectorySeparatorChar)))
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
						this.D(VH.A(54099));
						goto IL_03d7;
					}
				}
				this.C(ex2.Message);
				clsReporting.LogException(ex2);
				goto IL_03d7;
				IL_03d7:
				throw;
			}
			finally
			{
				range = null;
			}
		}
	}

	private void A(ChartObject A, Arrow B, int C, bool D, Range E, List<Range> F, string G, string H)
	{
		Range topLeftCell = A.TopLeftCell;
		int startPeriod = B.StartPeriod;
		int endPeriod = B.EndPeriod;
		checked
		{
			try
			{
				topLeftCell = ((_Worksheet)topLeftCell.Worksheet).get_Range((object)topLeftCell.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value));
				Range range;
				if (!D)
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
					range = topLeftCell.get_Resize((object)4, (object)4);
				}
				else
				{
					range = topLeftCell.get_Resize((object)5, (object)6);
				}
				topLeftCell = range;
				topLeftCell.Select();
				topLeftCell.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
				topLeftCell = (Range)this.m_A.Selection;
				topLeftCell.ClearFormats();
				Range range2 = topLeftCell;
				((Range)range2.Cells[1, 1]).Value2 = VH.A(53866);
				((Range)range2.Cells[2, 1]).Value2 = VH.A(54327);
				((Range)range2.Cells[3, 1]).Value2 = VH.A(54344);
				((Range)range2.Cells[4, 1]).Value2 = VH.A(54369);
				if (D)
				{
					((Range)range2.Cells[5, 1]).Value2 = VH.A(54386);
				}
				Range obj = (Range)range2.Cells[4, 2];
				obj.Value2 = YOffset;
				obj.Font.Color = C;
				string text = obj.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				_ = null;
				string text2 = default(string);
				if (D)
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
					Range obj2 = (Range)range2.Cells[5, 2];
					obj2.Value2 = YRise;
					obj2.Font.Color = C;
					text2 = obj2.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					_ = null;
				}
				if (F.Any())
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
					if (this.m_A.LineType == ArrowType.Flat)
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
						string text3 = VH.A(54403);
						int num = B.StartPeriod - 1;
						int num2 = B.EndPeriod - 1;
						for (int i = num; i <= num2; i++)
						{
							try
							{
								text3 = text3 + VH.A(54414) + H + F[i].get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(54423);
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								clsReporting.LogException(ex2);
								throw;
							}
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
						text3 = Strings.Left(text3, text3.Length - 1);
						text3 = text3 + VH.A(54428) + text;
						text3 = ExcelAddIn1.Formulas.Helpers.RemoveExtraneousSheetName(text3, G);
						((Range)range2.Cells[1, 2]).Formula = text3;
						int num3 = Conversions.ToInteger(range2.Columns.CountLarge);
						for (int j = 3; j <= num3; j++)
						{
							((Range)range2.Cells[1, j]).FormulaR1C1 = VH.A(54433);
						}
					}
					else
					{
						string text3 = VH.A(54448) + H + F[B.StartPeriod - 1].get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(54428) + text;
						text3 = ExcelAddIn1.Formulas.Helpers.RemoveExtraneousSheetName(text3, G);
						((Range)range2.Cells[1, 2]).Formula = text3;
						if (D)
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
							text3 = VH.A(54403);
							foreach (Range item in F.GetRange(B.StartPeriod - 1, B.EndPeriod - B.StartPeriod + 1))
							{
								text3 = text3 + VH.A(54414) + H + item.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(54423);
							}
							text3 = Strings.Left(text3, text3.Length - 1);
							text3 = text3 + VH.A(54428) + text + VH.A(54459) + text2;
							text3 = ExcelAddIn1.Formulas.Helpers.RemoveExtraneousSheetName(text3, G);
							int num4 = Conversions.ToInteger(Operators.SubtractObject(range2.Columns.CountLarge, 1));
							for (int k = 3; k <= num4; k++)
							{
								((Range)range2.Cells[1, k]).Formula = text3;
							}
						}
						else
						{
							((Range)range2.Cells[1, 3]).FormulaR1C1 = VH.A(54462);
						}
						text3 = VH.A(54448) + H + F[B.EndPeriod - 1].get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(54428) + text;
						text3 = ExcelAddIn1.Formulas.Helpers.RemoveExtraneousSheetName(text3, G);
						((Range)range2.Cells[1, RuntimeHelpers.GetObjectValue(range2.Columns.CountLarge)]).Formula = text3;
					}
				}
				else
				{
					((Range)range2.Cells[1, 2]).Formula = VH.A(48936) + Operators.SubtractObject(NewLateBinding.LateIndexGet(B.Series.Values, new object[1] { 1 }, null), YOffset).ToString() + VH.A(54459) + text;
					if (D)
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
						int num5 = Conversions.ToInteger(Operators.SubtractObject(range2.Columns.CountLarge, 1));
						for (int l = 3; l <= num5; l++)
						{
							((Range)range2.Cells[1, l]).Value2 = VH.A(48936) + Operators.SubtractObject(Operators.SubtractObject(NewLateBinding.LateIndexGet(B.Series.Values, new object[1] { l - 1 }, null), YOffset), YRise).ToString() + VH.A(54459) + text + VH.A(54459) + text2;
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
					else
					{
						((Range)range2.Cells[1, 3]).Value2 = VH.A(48936) + Operators.SubtractObject(Operators.SubtractObject(NewLateBinding.LateIndexGet(B.Series.Values, new object[1] { 2 }, null), YOffset), YRise).ToString() + VH.A(54459) + text;
					}
					((Range)range2.Cells[1, RuntimeHelpers.GetObjectValue(range2.Columns.CountLarge)]).Formula = VH.A(48936) + Operators.SubtractObject(NewLateBinding.LateIndexGet(B.Series.Values, new object[1] { Operators.SubtractObject(range2.Columns.CountLarge, 1) }, null), YOffset).ToString() + VH.A(54459) + text;
				}
				if (E != null)
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
					string text4;
					string text5;
					if (Operators.ConditionalCompareObjectEqual(E.Rows.CountLarge, 1, TextCompare: false))
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
						text4 = H + E.get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)endPeriod).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						text5 = VH.A(47375);
					}
					else
					{
						text4 = H + E.get_Resize((object)endPeriod, RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						text5 = VH.A(47401);
					}
					string text3 = VH.A(48936) + text5 + VH.A(39848) + text4 + VH.A(39904);
					text3 = ExcelAddIn1.Formulas.Helpers.RemoveExtraneousSheetName(text3, G);
					if (D)
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
						Range obj3 = (Range)range2.Cells[2, 2];
						obj3.Value2 = startPeriod;
						obj3.Font.Color = C;
						_ = null;
						Range obj4 = (Range)range2.Cells[2, 3];
						obj4.Value2 = startPeriod;
						obj4.Font.Color = C;
						_ = null;
						((Range)range2.Cells[2, 4]).FormulaR1C1 = VH.A(54507);
						((Range)range2.Cells[2, 5]).Formula = text3;
						((Range)range2.Cells[2, 6]).Formula = text3;
					}
					else
					{
						Range obj5 = (Range)range2.Cells[2, 2];
						obj5.Value2 = startPeriod;
						obj5.Font.Color = C;
						_ = null;
						((Range)range2.Cells[2, 3]).FormulaR1C1 = VH.A(54507);
						((Range)range2.Cells[2, 4]).Formula = text3;
					}
					E = null;
				}
				else
				{
					int num6 = Conversions.ToInteger(range2.Columns.CountLarge);
					for (int m = 2; m <= num6; m++)
					{
						((Range)range2.Cells[2, m]).Value2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(B.Series.XValues, new object[1] { m - 1 }, null));
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
				if (D)
				{
					this.A(B, (Range)range2.Cells[3, 4], C, F, H);
				}
				else
				{
					this.A(B, (Range)range2.Cells[3, 3], C, F, H);
				}
				range2.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(220, 220, 220));
				range2 = null;
				int num7 = Conversions.ToInteger(Operators.SubtractObject(topLeftCell.Columns.CountLarge, 1));
				Series series = B.Series;
				series.Name = VH.A(48936) + ((Range)topLeftCell.Cells[1, 1]).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
				series.Values = VH.A(48936) + ((Range)topLeftCell.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)0, (object)1).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)num7).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
				series.XValues = VH.A(48936) + ((Range)topLeftCell.Rows[2, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)0, (object)1).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)num7).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
				if (this.m_A.LabelPosition != CagrLabelPosition.None)
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
					DataLabels.LinkToRange((Microsoft.Office.Interop.Excel.DataLabels)series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), ((Range)topLeftCell.Rows[3, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)0, (object)1).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)num7));
					((Microsoft.Office.Interop.Excel.DataLabels)series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).AutoText = true;
				}
				series = null;
			}
			finally
			{
				topLeftCell = null;
			}
		}
	}

	private void A(Arrow A, Range B, int C, List<Range> D, string E)
	{
		Range growthRange = A.GrowthRange;
		checked
		{
			if (growthRange == null)
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
				int b = this.A();
				if ((float)A.GrowthRate == (float)A.A(ChartTotals, b))
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
					int index = A.StartPeriod - 1;
					int index2 = A.EndPeriod - 1;
					string strFormula = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(54556) + E + D[index2].get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(54569) + E + D[index].get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(54582) + b + VH.A(54591) + E, NewLateBinding.LateGet(D[index].Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)), VH.A(2826)), NewLateBinding.LateGet(D[index2].Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)), VH.A(39904)), VH.A(54608)));
					strFormula = ExcelAddIn1.Formulas.Helpers.RemoveExtraneousSheetName(strFormula, B.Worksheet.Name);
					NewLateBinding.LateSet(B, null, VH.A(1998), new object[1] { strFormula }, null, null);
				}
				else
				{
					B.Value2 = A.GrowthRate;
					B.Font.Color = C;
				}
			}
			else
			{
				string strFormula;
				if (growthRange.Worksheet == B.Worksheet)
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
					strFormula = VH.A(48936) + growthRange.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				else
				{
					strFormula = VH.A(54621) + growthRange.Worksheet.Name + VH.A(43343) + growthRange.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				NewLateBinding.LateSet(B, null, VH.A(1998), new object[1] { strFormula }, null, null);
			}
			try
			{
				B.NumberFormat = this.m_A.Format;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				this.D(VH.A(54626) + B.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(54723));
				ProjectData.ClearProjectError();
			}
			growthRange = null;
		}
	}

	private void AddArrow(object sender, RoutedEventArgs e)
	{
		A();
	}

	private void RemoveArrow(object sender, RoutedEventArgs e)
	{
		if (Forms.OkCancelMessage2(System.Windows.Window.GetWindow(this), VH.A(54901)) != DialogResult.OK)
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
			Arrow arrow = (Arrow)((System.Windows.Controls.Button)sender).DataContext;
			A(arrow.Series);
			GrowthArrows.Remove(arrow);
			arrow = null;
			this.m_A.ScreenUpdating = false;
			using (IEnumerator<Arrow> enumerator = GrowthArrows.GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					enumerator.Current.B(this.m_A.Format);
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_009c;
					}
					continue;
					end_IL_009c:
					break;
				}
			}
			this.m_A.ScreenUpdating = true;
			return;
		}
	}

	private void RecommendArrows(object sender, RoutedEventArgs e)
	{
		float num = 1f;
		int num2 = 1;
		List<int> list = new List<int>();
		list.Add(1);
		double num3 = A(1, 2);
		int count = ChartTotals.Count;
		checked
		{
			for (int i = 3; i <= count; i++)
			{
				double num4 = A(num2, i);
				if (!(num4 < num3 / (double)(1f + num)))
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
					if (!(num4 > num3 * (double)(1f + num)))
					{
						num3 = num4;
						continue;
					}
				}
				list.Add(i - 1);
				num2 = i - 1;
				num3 = A(num2, num2 + 1);
			}
			this.m_A.ScreenUpdating = false;
			try
			{
				F();
				int num5 = list.Count - 1;
				for (int j = 0; j <= num5; j++)
				{
					int b;
					if (j < list.Count - 1)
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
						b = list[j + 1];
					}
					else
					{
						b = ChartTotals.Count;
					}
					B(list[j], b);
				}
				while (true)
				{
					switch (6)
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
				this.m_A.ScreenUpdating = true;
				list = null;
			}
		}
	}

	private double A(int A, int B)
	{
		return checked(Math.Pow(ChartTotals[B - 1] / ChartTotals[A - 1], (double)this.A() / (double)(B - A))) - 1.0;
	}

	private void GrowthRate_GotFocus(object sender, RoutedEventArgs e)
	{
		Arrow obj = (Arrow)((System.Windows.Controls.TextBox)sender).DataContext;
		obj.GrowthRateText = obj.GrowthRate.ToString();
	}

	private void GrowthRate_LostFocus(object sender, RoutedEventArgs e)
	{
		System.Windows.Controls.TextBox textBox = (System.Windows.Controls.TextBox)sender;
		Arrow arrow = (Arrow)textBox.DataContext;
		if (double.TryParse(textBox.Text, out var result))
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
			arrow.A(result, this.m_A.Format);
		}
		else
		{
			D(VH.A(54938));
			textBox.Focus();
		}
		arrow = null;
		textBox = null;
	}

	private void ComputeCagrFromChartData(object sender, RoutedEventArgs e)
	{
		Arrow arrow = (Arrow)((System.Windows.Controls.Button)sender).DataContext;
		if (arrow.GrowthRateIsLinked)
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
			arrow.GrowthRateIsLinked = false;
		}
		arrow.A(ChartTotals, A());
		arrow.B(this.m_A.Format);
		arrow = null;
	}

	private void GrowthRateLinked(object sender, RoutedEventArgs e)
	{
		Arrow arrow = (Arrow)((System.Windows.Controls.CheckBox)sender).DataContext;
		string text = "";
		try
		{
			text = this.m_A.ActiveCell.get_Address((object)true, (object)true, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			Range range = (Range)this.m_A.InputBox(VH.A(54979), VH.A(40448), text, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), 8);
			if (range != null)
			{
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
					double result;
					if (Operators.ConditionalCompareObjectNotEqual(range.Cells.CountLarge, 1, TextCompare: false))
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
						D(VH.A(55066));
					}
					else if (double.TryParse(Conversions.ToString(range.Value2), out result))
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
						arrow.A(result, this.m_A.Format);
						arrow.GrowthRange = range;
						arrow.ManualEntryIsEnabled = false;
					}
					else
					{
						D(VH.A(55109));
					}
					range = null;
					break;
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		if (arrow.GrowthRange == null)
		{
			arrow.GrowthRateIsLinked = false;
		}
		Focus();
		arrow = null;
	}

	private void GrowthRateUnlinked(object sender, RoutedEventArgs e)
	{
		Arrow obj = (Arrow)((System.Windows.Controls.CheckBox)sender).DataContext;
		obj.GrowthRange = null;
		obj.ManualEntryIsEnabled = true;
	}

	private void PeriodChanged(object sender, SelectionChangedEventArgs e)
	{
		System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
		if (IsLoading)
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
			if (comboBox.SelectedIndex == -1)
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
				((Arrow)comboBox.DataContext).A(ChartTotals, A(), YOffset, YRise, this.m_A);
				comboBox = null;
				return;
			}
		}
	}

	private int A()
	{
		return 1;
	}

	private void LineStyleChanged(object sender, SelectionChangedEventArgs e)
	{
		this.m_A.ScreenUpdating = false;
		try
		{
			F();
			int selectedIndex = cbxStyle.SelectedIndex;
			bool a;
			if (selectedIndex != 0)
			{
				if (selectedIndex != 1)
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
					this.m_A.LineType = ArrowType.Flat;
					a = false;
					chkRotate.IsChecked = false;
					chkRotate.IsEnabled = false;
				}
				else
				{
					this.m_A.LineType = ArrowType.Angled;
					a = false;
					chkRotate.IsEnabled = true;
				}
			}
			else
			{
				this.m_A.LineType = ArrowType.Elbow;
				a = true;
				chkRotate.IsChecked = false;
				chkRotate.IsEnabled = false;
			}
			A(a);
			A().B(this.m_A.Format);
		}
		finally
		{
			this.m_A.ScreenUpdating = true;
		}
	}

	private void O(double A, double B)
	{
		IEnumerator<Arrow> enumerator = default(IEnumerator<Arrow>);
		try
		{
			enumerator = GrowthArrows.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Arrow current = enumerator.Current;
				object objectValue = RuntimeHelpers.GetObjectValue(current.Series.Values);
				int num = Conversions.ToInteger(NewLateBinding.LateGet(objectValue, null, VH.A(53246), new object[0], null, null, null));
				for (int i = 1; i <= num; i = checked(i + 1))
				{
					NewLateBinding.LateIndexSet(objectValue, new object[2]
					{
						i,
						Operators.AddObject(Operators.SubtractObject(NewLateBinding.LateIndexGet(objectValue, new object[1] { i }, null), A), B)
					}, null);
				}
				current.Series.Values = RuntimeHelpers.GetObjectValue(objectValue);
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
	}

	private void P(double A, double B)
	{
		if (this.m_A.LineType != ArrowType.Elbow)
		{
			return;
		}
		IEnumerator<Arrow> enumerator = default(IEnumerator<Arrow>);
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
			try
			{
				enumerator = GrowthArrows.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Arrow current = enumerator.Current;
					object objectValue = RuntimeHelpers.GetObjectValue(current.Series.Values);
					int num = Conversions.ToInteger(Operators.SubtractObject(NewLateBinding.LateGet(objectValue, null, VH.A(53246), new object[0], null, null, null), 1));
					for (int i = 2; i <= num; i = checked(i + 1))
					{
						NewLateBinding.LateIndexSet(objectValue, new object[2]
						{
							i,
							Operators.AddObject(Operators.SubtractObject(NewLateBinding.LateIndexGet(objectValue, new object[1] { i }, null), A), B)
						}, null);
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_00e4;
						}
						continue;
						end_IL_00e4:
						break;
					}
					current.Series.Values = RuntimeHelpers.GetObjectValue(objectValue);
				}
				while (true)
				{
					switch (2)
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
				if (enumerator != null)
				{
					while (true)
					{
						switch (6)
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
	}

	private void LineColorChanged(object sender, RoutedEventArgs e)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		wpfPalette val = new wpfPalette(false, (ColorRole)0);
		((System.Windows.Window)(object)val).Owner = System.Windows.Window.GetWindow(this);
		((System.Windows.Window)(object)val).Title = VH.A(55164);
		((System.Windows.Window)(object)val).ShowDialog();
		if (((System.Windows.Window)(object)val).DialogResult.HasValue && ((System.Windows.Window)(object)val).DialogResult.Value)
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
			if (val.SelectedColor != Colors.Transparent)
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
				btnColor.Foreground = new SolidColorBrush(val.SelectedColor);
				this.m_A.Color = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(val.SelectedColor.R, val.SelectedColor.G, val.SelectedColor.B));
				IEnumerator<Arrow> enumerator = default(IEnumerator<Arrow>);
				try
				{
					enumerator = GrowthArrows.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Arrow current;
						(current = enumerator.Current).Series.Format.Line.ForeColor.RGB = this.m_A.Color;
						if (current.Label != null)
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
							current.Label.Format.Line.ForeColor.RGB = this.m_A.Color;
							if (chkFill.IsChecked == true)
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
								current.Label.Format.Fill.ForeColor.RGB = this.m_A.Color;
							}
							if (chkMatch.IsChecked == true)
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
								current.Label.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = this.m_A.Color;
							}
						}
						current = null;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_0222;
						}
						continue;
						end_IL_0222:
						break;
					}
				}
				finally
				{
					if (enumerator != null)
					{
						while (true)
						{
							switch (6)
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
			else
			{
				D(VH.A(55189));
			}
		}
		val = null;
	}

	private void LineWeightChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		this.m_A.Weight = (float)numWeight.Value.Value;
		foreach (Arrow growthArrow in GrowthArrows)
		{
			growthArrow.Series.Format.Line.Weight = this.m_A.Weight;
			if (this.m_A.LabelPosition != CagrLabelPosition.Inline)
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
			growthArrow.Label.Format.Line.Weight = this.m_A.Weight;
		}
	}

	private void Q()
	{
		cbxShape.IsEnabled = false;
		txtFormat.IsEnabled = false;
		chkBold.IsEnabled = false;
		chkBorder.IsEnabled = false;
		chkFill.IsEnabled = false;
		chkMatch.IsEnabled = false;
		chkRotate.IsEnabled = false;
		cbxShape.IsEnabled = false;
		cbxShape.SelectionChanged -= ShapeTypeChanged;
		cbxShape.SelectedIndex = -1;
		cbxShape.SelectionChanged += ShapeTypeChanged;
	}

	private void R()
	{
		cbxShape.IsEnabled = true;
		W();
		if (!A())
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
			txtFormat.IsEnabled = true;
			txtFormat.Text = this.m_A.Format;
		}
		chkBold.IsEnabled = true;
		chkBorder.IsEnabled = true;
		chkBorder.IsChecked = true;
		chkFill.IsEnabled = true;
		bool? isChecked = chkFill.IsChecked;
		bool? obj;
		if (!isChecked.HasValue)
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
			obj = null;
		}
		else
		{
			obj = isChecked != true;
		}
		isChecked = obj;
		if (isChecked == true)
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
			chkFill.IsChecked = false;
			chkMatch.IsEnabled = true;
		}
		if (this.m_A.LineType != ArrowType.Angled)
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
			chkRotate.IsEnabled = true;
			return;
		}
	}

	private void S()
	{
		cbxShape.IsEnabled = false;
		cbxShape.SelectionChanged -= ShapeTypeChanged;
		cbxShape.SelectedIndex = -1;
		cbxShape.SelectionChanged += ShapeTypeChanged;
		if (!A())
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
			txtFormat.IsEnabled = true;
			txtFormat.Text = this.m_A.Format;
		}
		chkBold.IsEnabled = true;
		chkBorder.IsEnabled = false;
		chkBorder.IsChecked = false;
		chkFill.IsEnabled = false;
		chkFill.IsChecked = false;
		chkMatch.IsEnabled = true;
		if (this.m_A.LineType != ArrowType.Angled)
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
			chkRotate.IsEnabled = true;
			return;
		}
	}

	private void A(Arrow A, bool B)
	{
		DataLabel label = A.Label;
		if (B)
		{
			label.Format.TextFrame2.TextRange.Font.Fill.Visible = MsoTriState.msoTrue;
			D(A);
		}
		else
		{
			label.Format.TextFrame2.TextRange.Font.Fill.Visible = MsoTriState.msoFalse;
			label.Border.LineStyle = XlLineStyle.xlLineStyleNone;
			label.Fill.Visible = MsoTriState.msoFalse;
		}
		label = null;
	}

	private void LabelPositionChanged(object sender, SelectionChangedEventArgs e)
	{
		using IEnumerator<Arrow> enumerator = GrowthArrows.GetEnumerator();
		while (enumerator.MoveNext())
		{
			Arrow current = enumerator.Current;
			if (cbxPosition.SelectedIndex == 3)
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
				this.m_A.LabelPosition = CagrLabelPosition.None;
				A(current, B: false);
				Q();
				continue;
			}
			C(current);
			B(current);
			DataLabel label = current.Label;
			if (cbxPosition.SelectedIndex == 0)
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
				this.m_A.LabelPosition = CagrLabelPosition.Inline;
				A(current, B: true);
				R();
				label.Position = XlDataLabelPosition.xlLabelPositionCenter;
				LineFormat line = label.Format.Line;
				line.Visible = MsoTriState.msoTrue;
				line.Weight = this.m_A.Weight;
				line.ForeColor.RGB = this.m_A.Color;
				_ = null;
				label.Format.Fill.Visible = MsoTriState.msoTrue;
				label.Format.Fill.ForeColor.RGB = this.m_A;
				if (this.m_A.Rotate)
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
					if (this.m_A.LineType == ArrowType.Angled)
					{
						current.A();
					}
				}
			}
			else
			{
				if (cbxPosition.SelectedIndex == 1)
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
					this.m_A.LabelPosition = CagrLabelPosition.AboveLine;
				}
				else
				{
					this.m_A.LabelPosition = CagrLabelPosition.BelowLine;
				}
				if (this.m_A.LabelPosition == CagrLabelPosition.AboveLine)
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
					label.Position = XlDataLabelPosition.xlLabelPositionAbove;
				}
				else
				{
					label.Position = XlDataLabelPosition.xlLabelPositionBelow;
				}
				A(current, B: true);
				S();
				label.Format.AutoShapeType = MsoAutoShapeType.msoShapeRectangle;
				label.Format.Line.Visible = MsoTriState.msoFalse;
				label.Format.Fill.Visible = MsoTriState.msoFalse;
				if (this.m_A.LineType == ArrowType.Angled)
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
					if (this.m_A.Rotate)
					{
						current.A();
					}
				}
			}
			label = null;
		}
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

	private void ShapeTypeChanged(object sender, SelectionChangedEventArgs e)
	{
		Helpers.A(this.m_A, Chart);
		ref ArrowOptions a = ref this.m_A;
		switch (cbxShape.SelectedIndex)
		{
		case 0:
			a.Shape = MsoAutoShapeType.msoShapeRectangle;
			break;
		case 1:
			a.Shape = MsoAutoShapeType.msoShapeRoundedRectangle;
			break;
		case 2:
			a.Shape = MsoAutoShapeType.msoShapeOval;
			break;
		}
		IEnumerator<Arrow> enumerator = default(IEnumerator<Arrow>);
		try
		{
			enumerator = GrowthArrows.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Arrow current = enumerator.Current;
				if (current.Label == null)
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
					C(current);
				}
				current.Label.Format.AutoShapeType = this.m_A.Shape;
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
	}

	private void txtFormat_TextChanged(object sender, TextChangedEventArgs e)
	{
		IEnumerator<Arrow> enumerator = default(IEnumerator<Arrow>);
		try
		{
			enumerator = GrowthArrows.GetEnumerator();
			while (true)
			{
				if (enumerator.MoveNext())
				{
					Arrow current = enumerator.Current;
					try
					{
						current.Label.Text = Strings.Format(current.GrowthRate, txtFormat.Text);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
						break;
					}
					continue;
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
					switch (2)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
		Focus();
		txtFormat.Focus();
	}

	private void txtFormat_LostFocus(object sender, RoutedEventArgs e)
	{
		try
		{
			IEnumerator<Arrow> enumerator = default(IEnumerator<Arrow>);
			try
			{
				enumerator = GrowthArrows.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Arrow current = enumerator.Current;
					current.Label.Text = Strings.Format(current.GrowthRate, txtFormat.Text);
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
					break;
				}
			}
			finally
			{
				if (enumerator != null)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						enumerator.Dispose();
						break;
					}
				}
			}
			this.m_A.Format = txtFormat.Text;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			txtFormat.LostFocus -= txtFormat_LostFocus;
			C(VH.A(55246));
			txtFormat.LostFocus += txtFormat_LostFocus;
			txtFormat.Focus();
			txtFormat.SelectAll();
			ProjectData.ClearProjectError();
		}
	}

	private void chkBorder_CheckedChanged(object sender, RoutedEventArgs e)
	{
		this.m_A.LabelBorder = chkBorder.IsChecked.Value;
		IEnumerator<Arrow> enumerator = default(IEnumerator<Arrow>);
		try
		{
			enumerator = GrowthArrows.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Arrow current = enumerator.Current;
				if (current.Label == null)
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
					C(current);
				}
				LineFormat line = current.Label.Format.Line;
				if (chkBorder.IsChecked == true)
				{
					line.Visible = MsoTriState.msoTrue;
					line.Weight = this.m_A.Weight;
					line.ForeColor.RGB = this.m_A.Color;
				}
				else
				{
					line.Visible = MsoTriState.msoFalse;
				}
				line = null;
			}
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
	}

	private void chkFill_CheckedChanged(object sender, RoutedEventArgs e)
	{
		this.m_A.Invert = chkFill.IsChecked.Value;
		int rGB;
		int rGB2;
		if (chkFill.IsChecked == true)
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
			rGB = this.m_A.Color;
			rGB2 = this.m_A;
			System.Windows.Controls.CheckBox checkBox = chkMatch;
			checkBox.Checked -= chkMatch_CheckedChanged;
			checkBox.Unchecked -= chkMatch_CheckedChanged;
			checkBox.IsChecked = false;
			checkBox.IsEnabled = false;
			checkBox.Checked += chkMatch_CheckedChanged;
			checkBox.Unchecked += chkMatch_CheckedChanged;
			_ = null;
			this.m_A.MatchColor = false;
		}
		else
		{
			rGB = this.m_A;
			rGB2 = ((!this.m_A.MatchColor) ? B() : this.m_A.Color);
			chkMatch.IsEnabled = true;
		}
		using IEnumerator<Arrow> enumerator = GrowthArrows.GetEnumerator();
		while (enumerator.MoveNext())
		{
			ChartFormat format = enumerator.Current.Label.Format;
			format.Fill.Visible = MsoTriState.msoTrue;
			format.Fill.ForeColor.RGB = rGB;
			format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rGB2;
			_ = null;
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

	private void chkRotate_CheckedChanged(object sender, RoutedEventArgs e)
	{
		this.m_A.Rotate = chkRotate.IsChecked.Value;
		IEnumerator<Arrow> enumerator = default(IEnumerator<Arrow>);
		try
		{
			enumerator = GrowthArrows.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Arrow current = enumerator.Current;
				if (current.Label != null)
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
					bool? isChecked = chkRotate.IsChecked;
					bool? flag;
					if (!isChecked.HasValue)
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
						flag = isChecked;
					}
					else
					{
						flag = isChecked != true;
					}
					isChecked = flag;
					if (isChecked == true)
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
						current.Label.Orientation = 0;
					}
					else
					{
						current.A();
					}
				}
				current = null;
			}
			while (true)
			{
				switch (3)
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

	private void D(Arrow A)
	{
		this.m_A.MatchColor = chkMatch.IsChecked.Value;
		int a;
		if (chkMatch.IsChecked == true)
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
			a = A.Series.Format.Line.ForeColor.RGB;
		}
		else
		{
			try
			{
				a = B();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				a = 0;
				ProjectData.ClearProjectError();
			}
		}
		A.B(a);
	}

	private void chkBold_CheckedChanged(object sender, RoutedEventArgs e)
	{
		this.m_A.LabelBold = chkBold.IsChecked.Value;
		IEnumerator<Arrow> enumerator = default(IEnumerator<Arrow>);
		try
		{
			enumerator = GrowthArrows.GetEnumerator();
			while (enumerator.MoveNext())
			{
				enumerator.Current.A(chkBold.IsChecked.Value);
			}
		}
		finally
		{
			if (enumerator != null)
			{
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
					enumerator.Dispose();
					break;
				}
			}
		}
	}

	private void chkMatch_CheckedChanged(object sender, RoutedEventArgs e)
	{
		IEnumerator<Arrow> enumerator = default(IEnumerator<Arrow>);
		try
		{
			enumerator = GrowthArrows.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Arrow current = enumerator.Current;
				D(current);
			}
		}
		finally
		{
			if (enumerator != null)
			{
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
					enumerator.Dispose();
					break;
				}
			}
		}
	}

	private int B()
	{
		return Conversions.ToInteger(((Axis)Chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory)).TickLabels.Font.Color);
	}

	private void T()
	{
		ArrowType lineType = this.m_A.LineType;
		if (lineType != ArrowType.Elbow)
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
					if (lineType == ArrowType.Angled)
					{
						cbxStyle.SelectedIndex = 1;
						A(A: false);
						chkRotate.IsEnabled = true;
					}
					else
					{
						cbxStyle.SelectedIndex = 2;
						A(A: false);
						chkRotate.IsEnabled = false;
					}
					return;
				}
			}
		}
		cbxStyle.SelectedIndex = 0;
		A(A: true);
		chkRotate.IsEnabled = false;
	}

	private void U()
	{
		btnOk.Visibility = Visibility.Visible;
		btnCancel.Visibility = Visibility.Visible;
		btnClose.Visibility = Visibility.Collapsed;
		grpArrows.IsEnabled = true;
		txtFormat.IsEnabled = true;
		((UIElement)(object)numOffset).IsEnabled = true;
		((UIElement)(object)numHeight).IsEnabled = true;
		cbxStyle.IsEnabled = true;
	}

	private void V()
	{
		btnOk.Visibility = Visibility.Collapsed;
		btnCancel.Visibility = Visibility.Collapsed;
		btnClose.Visibility = Visibility.Visible;
		grpArrows.IsEnabled = false;
		txtFormat.IsEnabled = false;
		((UIElement)(object)numOffset).IsEnabled = false;
		A(A: false);
		cbxStyle.IsEnabled = false;
	}

	private void A(bool A)
	{
		MacNumericUpDown obj = numHeight;
		int isEnabled;
		if (!this.A())
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
			isEnabled = (A ? 1 : 0);
		}
		else
		{
			isEnabled = 0;
		}
		((UIElement)(object)obj).IsEnabled = (byte)isEnabled != 0;
		IsLoading = true;
		if (A)
		{
			YRise = ((Axis)Chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue)).MajorUnit;
		}
		else
		{
			YRise = 0.0;
		}
		IsLoading = false;
	}

	private void W()
	{
		MsoAutoShapeType shape = this.m_A.Shape;
		if (shape != MsoAutoShapeType.msoShapeRectangle)
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
					if (shape != MsoAutoShapeType.msoShapeRoundedRectangle)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								if (shape != MsoAutoShapeType.msoShapeOval)
								{
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
								}
								cbxShape.SelectedIndex = 2;
								return;
							}
						}
					}
					cbxShape.SelectedIndex = 1;
					return;
				}
			}
		}
		cbxShape.SelectedIndex = 0;
	}

	private ArrowOptions A()
	{
		ArrowOptions result = default(ArrowOptions);
		XmlNode xmlNode = KH.A.SettingsXml.DocumentElement.SelectSingleNode(VH.A(53001));
		result.LabelBold = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(55289)).InnerText);
		result.LabelBorder = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(55308)).InnerText);
		result.Color = clsColors.RGB2Ole(xmlNode.SelectSingleNode(VH.A(55331)).InnerText);
		result.Format = xmlNode.SelectSingleNode(VH.A(55342)).InnerText;
		result.Invert = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(55365)).InnerText);
		result.LabelPosition = (CagrLabelPosition)Conversions.ToInteger(xmlNode.SelectSingleNode(VH.A(55388)).InnerText);
		result.MatchColor = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(55415)).InnerText);
		result.Rotate = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(55446)).InnerText);
		result.Shape = (MsoAutoShapeType)Conversions.ToInteger(xmlNode.SelectSingleNode(VH.A(55469)).InnerText);
		result.LineType = (ArrowType)Conversions.ToInteger(xmlNode.SelectSingleNode(VH.A(55498)).InnerText);
		result.Weight = float.Parse(xmlNode.SelectSingleNode(VH.A(55517)).InnerText, CultureInfo.InvariantCulture);
		xmlNode = null;
		return result;
	}

	private void X()
	{
		XmlDocument settingsXml = KH.A.SettingsXml;
		XmlNode xmlNode = settingsXml.DocumentElement.SelectSingleNode(VH.A(53001));
		xmlNode.SelectSingleNode(VH.A(55289)).InnerText = (0 - (this.m_A.LabelBold ? 1 : 0)).ToString();
		xmlNode.SelectSingleNode(VH.A(55308)).InnerText = (0 - (this.m_A.LabelBorder ? 1 : 0)).ToString();
		xmlNode.SelectSingleNode(VH.A(55331)).InnerText = clsColors.Color2RGB(ColorTranslator.FromOle(this.m_A.Color));
		xmlNode.SelectSingleNode(VH.A(55342)).InnerText = this.m_A.Format;
		xmlNode.SelectSingleNode(VH.A(55365)).InnerText = Conversions.ToString(this.m_A.Invert);
		XmlNode xmlNode2 = xmlNode.SelectSingleNode(VH.A(55388));
		int labelPosition = (int)this.m_A.LabelPosition;
		xmlNode2.InnerText = labelPosition.ToString();
		xmlNode.SelectSingleNode(VH.A(55415)).InnerText = (0 - (this.m_A.MatchColor ? 1 : 0)).ToString();
		xmlNode.SelectSingleNode(VH.A(55446)).InnerText = (0 - (this.m_A.Rotate ? 1 : 0)).ToString();
		XmlNode xmlNode3 = xmlNode.SelectSingleNode(VH.A(55469));
		labelPosition = (int)this.m_A.Shape;
		xmlNode3.InnerText = labelPosition.ToString();
		XmlNode xmlNode4 = xmlNode.SelectSingleNode(VH.A(55498));
		labelPosition = (int)this.m_A.LineType;
		xmlNode4.InnerText = labelPosition.ToString();
		xmlNode.SelectSingleNode(VH.A(55517)).InnerText = this.m_A.Weight.ToString(CultureInfo.InvariantCulture);
		_ = null;
		KH.A.SaveSettings(settingsXml);
		settingsXml = null;
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
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_C = true;
			Uri resourceLocator = new Uri(VH.A(55538), UriKind.Relative);
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
		//IL_0085: Unknown result type (might be due to invalid IL or missing references)
		//IL_008f: Expected O, but got Unknown
		//IL_00a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ab: Expected O, but got Unknown
		//IL_00d9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e3: Expected O, but got Unknown
		if (connectionId == 2)
		{
			scroller = (ScrollViewer)target;
			return;
		}
		if (connectionId == 3)
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
					stkContents = (StackPanel)target;
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
					grpArrows = (System.Windows.Controls.GroupBox)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnAddArrow = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			btnRecommend = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 11)
		{
			numOffset = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					numHeight = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnColor = (System.Windows.Controls.Button)target;
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
					numWeight = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 15)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					cbxStyle = (System.Windows.Controls.ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					cbxPosition = (System.Windows.Controls.ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					cbxShape = (System.Windows.Controls.ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkBorder = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkFill = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
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
					chkRotate = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 21)
		{
			chkBold = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 22)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkMatch = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 23)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtFormat = (System.Windows.Controls.TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 24)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnClose = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 25)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnCancel = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 26)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnOk = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 27)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnDismiss = (System.Windows.Controls.Button)target;
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

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
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
			eventSetter.Event = Selector.SelectionChangedEvent;
			eventSetter.Handler = new SelectionChangedEventHandler(PeriodChanged);
			((System.Windows.Style)target).Setters.Add(eventSetter);
		}
		if (connectionId == 5)
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
			((System.Windows.Controls.Button)target).Click += RemoveArrow;
		}
		if (connectionId == 6)
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
			((System.Windows.Controls.TextBox)target).GotFocus += GrowthRate_GotFocus;
			((System.Windows.Controls.TextBox)target).LostFocus += GrowthRate_LostFocus;
		}
		if (connectionId == 7)
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
			((System.Windows.Controls.Button)target).Click += ComputeCagrFromChartData;
		}
		if (connectionId != 8)
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
			((System.Windows.Controls.CheckBox)target).Checked += GrowthRateLinked;
			((System.Windows.Controls.CheckBox)target).Unchecked += GrowthRateUnlinked;
			return;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
