using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Markup;
using A;
using Foo.Controls;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1.Shapes.Arrange;

[DesignerGenerated]
public sealed class wpfArrange : UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<int, int> A;

		public static Func<ShapeItem, bool> A;

		public static Func<Arrangement, bool> A;

		public static Func<Arrangement, bool> B;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal int A(int A)
		{
			return A;
		}

		[SpecialName]
		internal bool A(ShapeItem A)
		{
			return A.IsChecked;
		}

		[SpecialName]
		internal bool A(Arrangement A)
		{
			return A.IsChecked;
		}

		[SpecialName]
		internal bool B(Arrangement A)
		{
			return A is RectArrangement;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private ObservableCollection<ContainerOption> m_A;

	private ObservableCollection<Arrangement> m_A;

	private int m_A;

	[CompilerGenerated]
	private int m_B;

	private int m_C;

	private ShapeRange m_A;

	[CompilerGenerated]
	private Dictionary<Slide, List<ShapeItem>> m_A;

	[CompilerGenerated]
	private List<ShapeItem> m_A;

	[CompilerGenerated]
	private Action m_A;

	[CompilerGenerated]
	private RectangleF? m_A;

	[CompilerGenerated]
	private Container m_A;

	[CompilerGenerated]
	private Preferences m_A;

	private bool m_A;

	private bool m_B;

	[CompilerGenerated]
	private Slide m_A;

	private ObservableCollection<ShapeItem> m_A;

	private BackgroundWorker m_A;

	private ScrollViewer m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("grdLibrary")]
	private Grid m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnBack")]
	private Button m_A;

	[AccessedThroughProperty("btnSelect")]
	[CompilerGenerated]
	private Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("scroller")]
	private ScrollViewer m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkContainer")]
	private CheckBox m_A;

	[AccessedThroughProperty("lbxArrange")]
	[CompilerGenerated]
	private ListBox m_A;

	[AccessedThroughProperty("grpReorder")]
	[CompilerGenerated]
	private GroupBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkReorder")]
	private CheckBox m_B;

	[AccessedThroughProperty("lbxReorder")]
	[CompilerGenerated]
	private ListBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnMoveUp")]
	private Button m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnMoveDown")]
	private Button m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("grpPrefs")]
	private GroupBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPreferences")]
	private CheckBox m_C;

	[AccessedThroughProperty("numMaxShapes")]
	[CompilerGenerated]
	private MacNumericUpDown m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkScale")]
	private CheckBox m_D;

	[AccessedThroughProperty("cbxStretch")]
	[CompilerGenerated]
	private ComboBox m_A;

	[AccessedThroughProperty("chkCenter")]
	[CompilerGenerated]
	private CheckBox m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("numColumnSpacing")]
	private MacNumericUpDown m_B;

	[AccessedThroughProperty("numRowSpacing")]
	[CompilerGenerated]
	private MacNumericUpDown m_C;

	[AccessedThroughProperty("numPadding")]
	[CompilerGenerated]
	private MacNumericUpDown m_D;

	[AccessedThroughProperty("cbxCircleAlign")]
	[CompilerGenerated]
	private ComboBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("numCircleScale")]
	private MacNumericUpDown m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("numRotationAngle")]
	private MacNumericUpDown m_F;

	[CompilerGenerated]
	[AccessedThroughProperty("chkRotate")]
	private CheckBox m_F;

	[AccessedThroughProperty("chkBestFit")]
	[CompilerGenerated]
	private CheckBox m_G;

	private bool m_C;

	public ObservableCollection<ContainerOption> Containers
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(68962));
		}
	}

	public ObservableCollection<Arrangement> Arrangements
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(68983));
		}
	}

	public int ThumbnailHeight
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(69008));
		}
	}

	public int ThumbnailWidth
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
	}

	public int NumberOfShapes
	{
		get
		{
			return this.m_C;
		}
		set
		{
			this.m_C = value;
			A(AH.A(69039));
		}
	}

	private ShapeRange ShapeRange
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			if (value == null)
			{
				return;
			}
			IEnumerator enumerator = default(IEnumerator);
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
				NumberOfShapes = value.Count;
				ShapeItems = new List<ShapeItem>();
				try
				{
					enumerator = value.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Shape a = (Shape)enumerator.Current;
						ShapeItems.Add(new ShapeItem(a));
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
			}
		}
	}

	internal Dictionary<Slide, List<ShapeItem>> Slides
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

	private List<ShapeItem> ShapeItems
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

	private Action ParentCloseAction
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

	private RectangleF? TargetRect
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

	private Container Container
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

	private Preferences Prefs
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

	private Slide CurrentSlide
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

	public ObservableCollection<ShapeItem> ReorderShapes
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(69293));
		}
	}

	internal virtual Grid grdLibrary
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

	internal virtual Button btnBack
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
			RoutedEventHandler value2 = BackToParent;
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

	internal virtual Button btnSelect
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
			RoutedEventHandler value2 = ReselectShapes;
			Button button = this.m_B;
			if (button != null)
			{
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual ScrollViewer scroller
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

	internal virtual CheckBox chkContainer
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

	internal virtual ListBox lbxArrange
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

	internal virtual GroupBox grpReorder
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

	internal virtual CheckBox chkReorder
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
			RoutedEventHandler value2 = ReorderViewExpanded;
			RoutedEventHandler value3 = [SpecialName] (object a0, RoutedEventArgs a1) =>
			{
				J();
			};
			CheckBox checkBox = this.m_B;
			if (checkBox != null)
			{
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.m_B = value;
			checkBox = this.m_B;
			if (checkBox == null)
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
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual ListBox lbxReorder
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
			SizeChangedEventHandler value2 = ReorderListBoxSizeChanged;
			ListBox listBox = this.m_B;
			if (listBox != null)
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
				listBox.SizeChanged -= value2;
			}
			this.m_B = value;
			listBox = this.m_B;
			if (listBox != null)
			{
				listBox.SizeChanged += value2;
			}
		}
	}

	internal virtual Button btnMoveUp
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

	internal virtual Button btnMoveDown
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

	internal virtual GroupBox grpPrefs
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

	internal virtual CheckBox chkPreferences
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
			RoutedEventHandler value2 = PrefsOpened;
			CheckBox checkBox = this.m_C;
			if (checkBox != null)
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
				checkBox.Checked -= value2;
			}
			this.m_C = value;
			checkBox = this.m_C;
			if (checkBox != null)
			{
				checkBox.Checked += value2;
			}
		}
	}

	internal virtual MacNumericUpDown numMaxShapes
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

	internal virtual CheckBox chkScale
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

	internal virtual ComboBox cbxStretch
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

	internal virtual CheckBox chkCenter
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

	internal virtual MacNumericUpDown numColumnSpacing
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

	internal virtual MacNumericUpDown numRowSpacing
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

	internal virtual MacNumericUpDown numPadding
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

	internal virtual ComboBox cbxCircleAlign
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

	internal virtual MacNumericUpDown numCircleScale
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

	internal virtual MacNumericUpDown numRotationAngle
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

	internal virtual CheckBox chkRotate
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

	internal virtual CheckBox chkBestFit
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
		}
	}

	public wpfArrange(RectangleF? rect, Action actClose)
	{
		//IL_0362: Unknown result type (might be due to invalid IL or missing references)
		//IL_036c: Expected O, but got Unknown
		//IL_037b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0385: Expected O, but got Unknown
		//IL_0394: Unknown result type (might be due to invalid IL or missing references)
		//IL_039e: Expected O, but got Unknown
		//IL_03ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_03b7: Expected O, but got Unknown
		//IL_03df: Unknown result type (might be due to invalid IL or missing references)
		//IL_03e9: Expected O, but got Unknown
		//IL_03f6: Unknown result type (might be due to invalid IL or missing references)
		//IL_0400: Expected O, but got Unknown
		base.Unloaded += ArrangeClosed;
		this.m_B = 160;
		this.m_B = false;
		CurrentSlide = null;
		this.m_A = null;
		InitializeComponent();
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		ParentCloseAction = actClose;
		TargetRect = rect;
		Prefs = new Preferences();
		Containers = new ObservableCollection<ContainerOption>();
		Arrangements = new ObservableCollection<Arrangement>();
		ShapeItems = new List<ShapeItem>();
		Grid grid = grdLibrary;
		int visibility;
		if (actClose != null)
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
			visibility = 0;
		}
		else
		{
			visibility = 2;
		}
		grid.Visibility = (Visibility)visibility;
		Master slideMaster = application.ActivePresentation.SlideMaster;
		ThumbnailHeight = checked((int)Math.Round((float)ThumbnailWidth * slideMaster.Height / slideMaster.Width));
		float height = slideMaster.Height;
		slideMaster = null;
		A(A: false);
		K();
		ProcessSelection(application.ActiveWindow.Selection);
		A(numPadding);
		A(numRowSpacing);
		A(numColumnSpacing);
		chkPreferences.IsChecked = false;
		Preferences prefs = Prefs;
		chkCenter.IsChecked = prefs.CenterShapes;
		numMaxShapes.Value = prefs.MaxShapesPerSlide;
		numPadding.Value = prefs.ConvertFromPoints(Math.Min(prefs.ContainerPadding, height / 2f));
		numPadding.Maximum = prefs.ConvertFromPoints((float)Math.Floor(height / 2f));
		numRowSpacing.Value = prefs.ConvertFromPoints(prefs.MinRowSpacing);
		numColumnSpacing.Value = prefs.ConvertFromPoints(prefs.MinColumnSpacing);
		chkScale.IsChecked = prefs.ScaleMode == ScaleMode.UniformArea;
		cbxStretch.SelectedIndex = (int)prefs.StretchMode;
		chkBestFit.IsChecked = prefs.ReorderBestFit;
		cbxCircleAlign.SelectedIndex = (int)prefs.CircleAlign;
		numCircleScale.Value = prefs.CircleScale;
		numRotationAngle.Value = prefs.RotationAngle;
		prefs = null;
		G();
		chkCenter.Checked += CenterChanged;
		chkCenter.Unchecked += CenterChanged;
		chkScale.Checked += ScaleModeChanged;
		chkScale.Unchecked += ScaleModeChanged;
		cbxStretch.SelectionChanged += StretchModeChanged;
		chkBestFit.Checked += BestFitChanged;
		chkBestFit.Unchecked += BestFitChanged;
		numMaxShapes.ValueChanged += new MacRangeBaseValueChangedHandler(MaxShapesPerSlideChanged);
		numPadding.ValueChanged += new MacRangeBaseValueChangedHandler(ContainerPaddingChanged);
		numRowSpacing.ValueChanged += new MacRangeBaseValueChangedHandler(MinRowSpacingChanged);
		numColumnSpacing.ValueChanged += new MacRangeBaseValueChangedHandler(MinColumnSpacingChanged);
		cbxCircleAlign.SelectionChanged += CircleAlignChanged;
		numCircleScale.ValueChanged += new MacRangeBaseValueChangedHandler(CircleScaleChanged);
		numRotationAngle.ValueChanged += new MacRangeBaseValueChangedHandler(RotationAngleChanged);
		chkRotate.Checked += RotateShapesChanged;
		chkRotate.Unchecked += RotateShapesChanged;
		application = null;
		H();
	}

	private void A(string A)
	{
		this.m_A?.Invoke(this, new PropertyChangedEventArgs(A));
	}

	private void A(MacNumericUpDown A)
	{
		MacNumericUpDown val = A;
		if (RegionInfo.CurrentRegion.IsMetric)
		{
			val.CustomUnit = AH.A(8238);
			val.SmallChange = 1.0;
			val.LargeChange = 10.0;
			val.NumberDecimalDigits = 0;
		}
		else
		{
			val.CustomUnit = AH.A(69068);
			val.SmallChange = 0.25;
			val.LargeChange = 1.0;
			val.NumberDecimalDigits = 2;
		}
		val = null;
	}

	private void A(Slide A)
	{
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = NG.A.Application.ActivePresentation;
		ContainerCanvas canvas = new ContainerCanvas
		{
			Height = ThumbnailHeight,
			Width = ThumbnailWidth,
			Scale = (float)ThumbnailWidth / activePresentation.SlideMaster.Width
		};
		ObservableCollection<ContainerOption> containers = Containers;
		containers.Clear();
		if (TargetRect.HasValue)
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
			containers.Add(new ContainerOption(AH.A(69073), canvas, TargetRect.Value));
		}
		containers.Add(new ContainerOption(AH.A(69098)));
		Settings settings = new Settings((Microsoft.Office.Interop.PowerPoint.Presentation)A.Parent);
		if (settings.SlideMargins.HasValue)
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
			containers.Add(new ContainerOption(AH.A(69109), canvas, settings.SlideMargins.Value));
		}
		else
		{
			try
			{
				containers.Add(new ContainerOption(AH.A(69109), canvas, Helpers.GetBodyPlaceholder(activePresentation)));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		settings = null;
		foreach (Shape placeholder in A.CustomLayout.Shapes.Placeholders)
		{
			PpPlaceholderType type = placeholder.PlaceholderFormat.Type;
			if (type <= PpPlaceholderType.ppPlaceholderBitmap)
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
				if (type != PpPlaceholderType.ppPlaceholderMixed)
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
					if ((uint)(type - 7) > 2u)
					{
						continue;
					}
				}
			}
			else if (type != PpPlaceholderType.ppPlaceholderTable)
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
				if (type != PpPlaceholderType.ppPlaceholderPicture)
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
			}
			containers.Add(new ContainerOption(AH.A(69136), canvas, placeholder));
		}
		this.m_A = true;
		containers[0].IsChecked = true;
		this.m_A = false;
		containers = null;
		this.A(A: true);
		activePresentation = null;
	}

	private void A(bool A)
	{
		chkContainer.IsChecked = A;
		chkContainer.IsEnabled = A;
	}

	private void A()
	{
		int count = ShapeItems.Count;
		Arrangements.Clear();
		int num = Math.Min(count, Prefs.MaxShapesPerSlide);
		checked
		{
			int num2 = (int)Math.Ceiling((double)count / (double)Prefs.MaxShapesPerSlide);
			int num3 = (int)Math.Ceiling((double)num / 5.0);
			int num4 = Math.Min(12, num);
			int num5 = num3;
			int num6 = num4;
			while (num6 >= num5)
			{
				List<RowItem> list = new List<RowItem>();
				int num7 = num;
				while (num7 > 0)
				{
					int num8 = Math.Min(num6, num7);
					list.Add(new RowItem(num8, Prefs.CenterShapes));
					num7 -= num8;
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
					if (list.Count > 1)
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
						if ((double)list.Last().Columns.Count < 0.33 * (double)list[list.Count - 2].Columns.Count)
						{
							goto IL_0149;
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
					Arrangements.Add(new RectArrangement(list, num6, num2));
					list = null;
					goto IL_0149;
					IL_0149:
					num6 += -1;
					break;
				}
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				if (num2 == 1)
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
					if (count <= 15)
					{
						if (count != 6)
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
							if (count != 10)
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
								if (count != 15)
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
									goto IL_020b;
								}
							}
						}
					}
					else if (count <= 28)
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
						if (count != 21)
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
							if (count != 28)
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
								goto IL_020b;
							}
						}
					}
					else if (count != 36)
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
						if (count != 45)
						{
							goto IL_020b;
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
					Arrangements.Add(new PyramidArrangement(count, 1));
				}
				goto IL_020b;
				IL_020b:
				Arrangements.Add(new CircleArrangement(num2));
				return;
			}
		}
	}

	public void ArrangementChecked(object sender, RoutedEventArgs e)
	{
		A((Arrangement)((RadioButton)sender).DataContext);
		chkReorder.IsEnabled = true;
	}

	private void B()
	{
		IEnumerator<Arrangement> enumerator = default(IEnumerator<Arrangement>);
		try
		{
			enumerator = Arrangements.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Arrangement current = enumerator.Current;
				if (current.IsChecked)
				{
					A(current);
					break;
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

	private void A(Arrangement A)
	{
		I();
		try
		{
			if (Slides.Count != 1)
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
				C();
			}
			if (A.SlideCount > 1)
			{
				this.A(A.SlideCount);
			}
			if (Prefs.ReorderBestFit)
			{
				using Dictionary<Slide, List<ShapeItem>>.Enumerator enumerator = Slides.GetEnumerator();
				while (enumerator.MoveNext())
				{
					A.B(enumerator.Current.Value, Container, Prefs);
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_0103;
					}
					continue;
					end_IL_0103:
					break;
				}
			}
			else
			{
				double D = 0.0;
				using Dictionary<Slide, List<ShapeItem>>.Enumerator enumerator2 = Slides.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					A.A(enumerator2.Current.Value, Container, Prefs, ref D);
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_00a7;
					}
					continue;
					end_IL_00a7:
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
		if (Slides.Count > 1)
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
			this.m_B = true;
		}
		H();
		if (Slides.Count == 1)
		{
			btnSelect.IsEnabled = true;
			chkReorder.IsEnabled = true;
		}
		else
		{
			btnSelect.IsEnabled = false;
			K();
		}
	}

	private void C()
	{
		List<ShapeItem> list = new List<ShapeItem>();
		Dictionary<Slide, List<ShapeItem>> slides = Slides;
		KeyValuePair<Slide, List<ShapeItem>> keyValuePair = slides.ElementAt(0);
		Slide key = keyValuePair.Key;
		list.AddRange(keyValuePair.Value.ToList());
		keyValuePair = default(KeyValuePair<Slide, List<ShapeItem>>);
		checked
		{
			int num = slides.Count - 1;
			IEnumerator enumerator2 = default(IEnumerator);
			for (int i = 1; i <= num; i++)
			{
				KeyValuePair<Slide, List<ShapeItem>> keyValuePair2 = slides.ElementAt(i);
				List<int> list2 = new List<int>();
				using (List<ShapeItem>.Enumerator enumerator = keyValuePair2.Value.GetEnumerator())
				{
					while (enumerator.MoveNext())
					{
						ShapeItem current = enumerator.Current;
						try
						{
							list2.Add(Helpers.A(keyValuePair2.Key, current.Shape));
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							D();
							ProjectData.ClearProjectError();
						}
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
				keyValuePair2.Key.Shapes.Range(list2.ToArray()).Cut();
				try
				{
					enumerator2 = key.Shapes.Paste().GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Shape a = (Shape)enumerator2.Current;
						list.Add(new ShapeItem(a));
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_0140;
						}
						continue;
						end_IL_0140:
						break;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
				keyValuePair2.Key.Delete();
				keyValuePair2 = default(KeyValuePair<Slide, List<ShapeItem>>);
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				slides.Clear();
				slides.Add(key, list);
				slides = null;
				List<int> list2;
				try
				{
					if (NG.A.Application.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							list2 = new List<int>();
							using (List<ShapeItem>.Enumerator enumerator3 = list.GetEnumerator())
							{
								while (enumerator3.MoveNext())
								{
									ShapeItem current2 = enumerator3.Current;
									list2.Add(Helpers.A(key, current2.Shape));
								}
								while (true)
								{
									switch (4)
									{
									case 0:
										break;
									default:
										goto end_IL_0215;
									}
									continue;
									end_IL_0215:
									break;
								}
							}
							ShapeRange = key.Shapes.Range(list2.ToArray());
							ShapeRange.Select();
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
				key = null;
				list = null;
				list2 = null;
				return;
			}
		}
	}

	private void A(int A)
	{
		Slide key = Slides.ElementAt(0).Key;
		List<Slide> list = new List<Slide>();
		List<int> list2 = new List<int>();
		checked
		{
			try
			{
				using (List<ShapeItem>.Enumerator enumerator = Slides.ElementAt(0).Value.GetEnumerator())
				{
					while (enumerator.MoveNext())
					{
						ShapeItem current = enumerator.Current;
						try
						{
							list2.Add(Helpers.A(key, current.Shape));
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							D();
							ProjectData.ClearProjectError();
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
				List<int> source = list2;
				Func<int, int> keySelector;
				if (_Closure_0024__.A == null)
				{
					keySelector = (_Closure_0024__.A = [SpecialName] (int result) => result);
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
					keySelector = _Closure_0024__.A;
				}
				list2 = source.OrderBy(keySelector).ToList();
				new ComAwareEventInfo(typeof(EApplication_Event), AH.A(58943)).AddEventHandler(key.Application, new EApplication_PresentationNewSlideEventHandler(Create.Application_PresentationNewSlide));
				int num = A - 1;
				for (int num2 = 1; num2 <= num; num2++)
				{
					list.Add(key.Duplicate()[1]);
				}
				new ComAwareEventInfo(typeof(EApplication_Event), AH.A(58943)).RemoveEventHandler(key.Application, new EApplication_PresentationNewSlideEventHandler(Create.Application_PresentationNewSlide));
				list.Add(key);
				int maxShapesPerSlide = Prefs.MaxShapesPerSlide;
				int num3 = 1;
				int num4 = maxShapesPerSlide;
				Slides.Clear();
				ShapeItems.Clear();
				for (int num5 = list.Count - 1; num5 >= 0; num5 += -1)
				{
					List<ShapeItem> list3 = new List<ShapeItem>();
					Shape shape;
					for (int num6 = list2.Count; num6 >= 1; shape = null, num6 += -1)
					{
						shape = list[num5].Shapes.Range(list2[num6 - 1])[1];
						if (num6 >= num3)
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
							if (num6 <= num4)
							{
								list3.Add(new ShapeItem(shape));
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
						}
						shape.Delete();
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_0260;
						}
						continue;
						end_IL_0260:
						break;
					}
					list3.Reverse();
					Slides.Add(list[num5], list3);
					ShapeItems.AddRange(list3);
					list3 = null;
					num3 += maxShapesPerSlide;
					num4 += maxShapesPerSlide;
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
				list = null;
				list2 = null;
				key = null;
			}
		}
	}

	private void D()
	{
		throw new Exception(AH.A(69159));
	}

	private void ContainerChecked(object sender, RoutedEventArgs e)
	{
		if (!this.m_A)
		{
			A((ContainerOption)((RadioButton)sender).DataContext);
			B();
		}
	}

	private void A(ContainerOption A)
	{
		Container = new Container(Prefs, A.Rectangle);
	}

	private void E()
	{
		foreach (ContainerOption container in Containers)
		{
			if (container.IsChecked)
			{
				A(container);
				break;
			}
		}
	}

	private void PrefsOpened(object sender, RoutedEventArgs e)
	{
		chkReorder.IsChecked = false;
		grpPrefs.BringIntoView();
	}

	private void MaxShapesPerSlideChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		Prefs.SaveMaxShapesPerSlide(checked((int)Math.Round(numMaxShapes.Value.Value)));
		A();
	}

	private void ContainerPaddingChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		Prefs.SaveContainerPadding((float)((MacNumericUpDown)sender).Value.Value);
		E();
		B();
	}

	private void CenterChanged(object sender, RoutedEventArgs e)
	{
		Prefs.SaveCenter(chkCenter.IsChecked.Value);
		foreach (Arrangement arrangement in Arrangements)
		{
			if (!(arrangement is RectArrangement))
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
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			RowItem rowItem = ((RectArrangement)arrangement).Rows.Last();
			int horizontalAlignment;
			if (chkCenter.IsChecked != true)
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
				horizontalAlignment = 0;
			}
			else
			{
				horizontalAlignment = 1;
			}
			rowItem.HorizontalAlignment = (HorizontalAlignment)horizontalAlignment;
		}
		if (!chkCenter.IsEnabled)
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
			B();
			return;
		}
	}

	private void ScaleModeChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked = chkScale.IsChecked;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
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
			F();
		}
		Preferences prefs = Prefs;
		int scale;
		if (chkScale.IsChecked != true)
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
			scale = 0;
		}
		else
		{
			scale = 1;
		}
		prefs.SaveScaleMode((ScaleMode)scale);
		B();
	}

	private void StretchModeChanged(object sender, SelectionChangedEventArgs e)
	{
		G();
		Prefs.SaveStretchMode((Stretch)cbxStretch.SelectedIndex);
		B();
	}

	private void MinRowSpacingChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		//IL_0009: Unknown result type (might be due to invalid IL or missing references)
		Prefs.SaveMinRowSpacing((float)((MacNumericUpDown)sender).Value.Value);
		B();
	}

	private void MinColumnSpacingChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		Prefs.SaveMinColumnSpacing((float)((MacNumericUpDown)sender).Value.Value);
		B();
	}

	private void CircleAlignChanged(object sender, SelectionChangedEventArgs e)
	{
		Prefs.SaveCircleAlign((CircleAlign)cbxCircleAlign.SelectedIndex);
		B();
	}

	private void CircleScaleChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		//IL_0009: Unknown result type (might be due to invalid IL or missing references)
		Prefs.SaveCircleScale(checked((int)Math.Round(((MacNumericUpDown)sender).Value.Value)));
		B();
	}

	private void RotationAngleChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		//IL_0009: Unknown result type (might be due to invalid IL or missing references)
		Prefs.SaveRotationAngle(checked((int)Math.Round(((MacNumericUpDown)sender).Value.Value)));
		B();
	}

	private void RotateShapesChanged(object sender, RoutedEventArgs e)
	{
		Prefs.SaveRotateShapes(chkRotate.IsChecked.Value);
		B();
	}

	private void BestFitChanged(object sender, RoutedEventArgs e)
	{
		Prefs.SaveBestFit(chkBestFit.IsChecked.Value);
		A();
	}

	private void F()
	{
		using List<ShapeItem>.Enumerator enumerator = ShapeItems.GetEnumerator();
		while (enumerator.MoveNext())
		{
			enumerator.Current.A();
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

	private void G()
	{
		if (cbxStretch.SelectedIndex != 0)
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
			if (cbxStretch.SelectedIndex != 1)
			{
				chkCenter.IsEnabled = true;
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
		}
		chkCenter.IsEnabled = false;
		chkCenter.IsChecked = false;
	}

	private void H()
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(NG.A.Application, new EApplication_WindowSelectionChangeEventHandler(A));
	}

	private void I()
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(NG.A.Application, new EApplication_WindowSelectionChangeEventHandler(A));
	}

	private void A(Selection A)
	{
		if (this.m_B)
		{
			this.m_B = false;
		}
		else
		{
			ProcessSelection(A);
		}
	}

	public void ProcessSelection(Selection sel)
	{
		Slides = new Dictionary<Slide, List<ShapeItem>>();
		Arrangements.Clear();
		try
		{
			PpSelectionType type = sel.Type;
			if (type != PpSelectionType.ppSelectionSlides)
			{
				IEnumerator enumerator = default(IEnumerator);
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
					if (type == PpSelectionType.ppSelectionShapes)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							List<ShapeItem> list = new List<ShapeItem>();
							ShapeRange = sel.ShapeRange;
							try
							{
								enumerator = sel.ShapeRange.GetEnumerator();
								while (enumerator.MoveNext())
								{
									Shape a = (Shape)enumerator.Current;
									list.Add(new ShapeItem(a));
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
							Slides.Add(sel.SlideRange[1], list);
							list = null;
							if (Containers.Count == 0)
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
								A(sel.SlideRange[1]);
							}
							A();
							break;
						}
					}
					else if (sel.SlideRange != null)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							if (CurrentSlide != sel.SlideRange[1])
							{
								A(sel.SlideRange[1]);
							}
							break;
						}
					}
					else
					{
						Containers.Clear();
						A(A: false);
					}
					break;
				}
			}
			else if (CurrentSlide != sel.SlideRange[1])
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					A(sel.SlideRange[1]);
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Containers.Clear();
			A(A: false);
			ProjectData.ClearProjectError();
		}
		K();
		try
		{
			CurrentSlide = sel.SlideRange[1];
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			CurrentSlide = null;
			ProjectData.ClearProjectError();
		}
	}

	private void ReorderViewExpanded(object sender, RoutedEventArgs e)
	{
		chkPreferences.IsChecked = false;
		grpReorder.BringIntoView();
		ReorderShapes = new ObservableCollection<ShapeItem>();
		using (List<ShapeItem>.Enumerator enumerator = Slides[CurrentSlide].GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				ShapeItem current = enumerator.Current;
				ReorderShapes.Add(current);
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
				break;
			}
		}
		if (this.m_A == null)
		{
			this.m_A = new BackgroundWorker();
			BackgroundWorker a = this.m_A;
			a.WorkerSupportsCancellation = true;
			a.WorkerReportsProgress = false;
			a.DoWork += bgw_DoWork;
			a.RunWorkerCompleted += bgw_RunWorkerCompleted;
			_ = null;
		}
		this.m_A.RunWorkerAsync();
	}

	private void J()
	{
		B(A: false);
		IEnumerator<ShapeItem> enumerator = default(IEnumerator<ShapeItem>);
		try
		{
			enumerator = ReorderShapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				enumerator.Current.IsChecked = false;
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

	private void ReorderListBoxSizeChanged(object sender, SizeChangedEventArgs e)
	{
		if (!e.HeightChanged)
		{
			return;
		}
		if (this.m_A == null)
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
			this.m_A = (ScrollViewer)Forms.GetScrollViewer((DependencyObject)lbxReorder);
		}
		ScrollViewer a = this.m_A;
		Thickness padding;
		if (this.m_A.ComputedVerticalScrollBarVisibility != Visibility.Visible)
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
			padding = new Thickness(0.0);
		}
		else
		{
			padding = new Thickness(0.0, 0.0, 3.0, 0.0);
		}
		a.Padding = padding;
	}

	private void bgw_DoWork(object sender, DoWorkEventArgs e)
	{
		foreach (ShapeItem reorderShape in ReorderShapes)
		{
			if (this.m_A.CancellationPending)
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
						return;
					}
				}
			}
			reorderShape.B();
		}
	}

	private void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
	}

	private void K()
	{
		chkReorder.IsChecked = false;
		chkReorder.IsEnabled = false;
	}

	private void B(bool A)
	{
		btnMoveUp.IsEnabled = A;
		btnMoveDown.IsEnabled = A;
	}

	private void ReorderShapeChecked(object sender, RoutedEventArgs e)
	{
		B(A: true);
	}

	private void MoveShapeUp(object sender, RoutedEventArgs e)
	{
		ShapeItem shapeItem = A();
		if (shapeItem == null)
		{
			return;
		}
		checked
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
				int num = ReorderShapes.IndexOf(shapeItem);
				if (num > 0)
				{
					Slides[CurrentSlide].Reverse(num - 1, 2);
					ReorderShapes.Move(num, num - 1);
				}
				else
				{
					List<ShapeItem> list = Slides[CurrentSlide];
					ShapeItem item = list[0];
					list.RemoveAt(0);
					list.Insert(list.Count, item);
					item = null;
					_ = null;
					ReorderShapes.Move(num, ReorderShapes.Count - 1);
				}
				lbxReorder.ScrollIntoView(shapeItem);
				B();
				shapeItem = null;
				return;
			}
		}
	}

	private void MoveShapeDown(object sender, RoutedEventArgs e)
	{
		ShapeItem shapeItem = A();
		if (shapeItem == null)
		{
			return;
		}
		checked
		{
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
				int num = ReorderShapes.IndexOf(shapeItem);
				if (num < ReorderShapes.Count - 1)
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
					Slides[CurrentSlide].Reverse(num, 2);
					ReorderShapes.Move(num, num + 1);
				}
				else
				{
					List<ShapeItem> list = Slides[CurrentSlide];
					ShapeItem item = list[list.Count - 1];
					list.RemoveAt(list.Count - 1);
					list.Insert(0, item);
					item = null;
					_ = null;
					ReorderShapes.Move(num, 0);
				}
				lbxReorder.ScrollIntoView(shapeItem);
				B();
				shapeItem = null;
				return;
			}
		}
	}

	private ShapeItem A()
	{
		ShapeItem result;
		try
		{
			ObservableCollection<ShapeItem> reorderShapes = ReorderShapes;
			Func<ShapeItem, bool> predicate;
			if (_Closure_0024__.A == null)
			{
				predicate = (_Closure_0024__.A = [SpecialName] (ShapeItem A) => A.IsChecked);
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
				predicate = _Closure_0024__.A;
			}
			result = reorderShapes.Where(predicate).ElementAtOrDefault(0);
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

	private void DeleteShape(object sender, RoutedEventArgs e)
	{
		if (ReorderShapes.Count == 1)
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
					Forms.WarningMessage(AH.A(69320));
					return;
				}
			}
		}
		Arrangement arrangement = null;
		IEnumerator<Arrangement> enumerator = default(IEnumerator<Arrangement>);
		try
		{
			enumerator = Arrangements.GetEnumerator();
			while (true)
			{
				if (enumerator.MoveNext())
				{
					Arrangement current = enumerator.Current;
					if (!current.IsChecked)
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
						arrangement = current;
						break;
					}
					break;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_0072;
					}
					continue;
					end_IL_0072:
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
					switch (3)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
		ShapeItem shapeItem = (ShapeItem)((Button)sender).DataContext;
		Slides[CurrentSlide].Remove(shapeItem);
		ReorderShapes.Remove(shapeItem);
		B(A: false);
		try
		{
			NewLateBinding.LateSetComplex(shapeItem.Shape, null, AH.A(69417), new object[1] { false }, null, null, OptimisticSet: false, RValueBase: true);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		I();
		this.m_B = true;
		try
		{
			shapeItem.Shape.Delete();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		ShapeRange = NG.A.Application.ActiveWindow.Selection.ShapeRange;
		H();
		shapeItem = null;
		A();
		if (arrangement != null)
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
			if (arrangement is RectArrangement)
			{
				RectArrangement rectArrangement = (RectArrangement)arrangement;
				if (rectArrangement.Rows.Count == 1)
				{
					B(1);
				}
				else if (rectArrangement.Rows.Count == checked(ReorderShapes.Count + 1))
				{
					B(ReorderShapes.Count);
				}
				else
				{
					B(rectArrangement.Rows.Count);
				}
				rectArrangement = null;
			}
			else if (arrangement is CircleArrangement)
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
				IEnumerator<Arrangement> enumerator2 = default(IEnumerator<Arrangement>);
				try
				{
					enumerator2 = Arrangements.GetEnumerator();
					while (true)
					{
						if (enumerator2.MoveNext())
						{
							Arrangement current2 = enumerator2.Current;
							if (!(current2 is CircleArrangement))
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
								current2.IsChecked = true;
								break;
							}
							break;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_025c;
							}
							continue;
							end_IL_025c:
							break;
						}
						break;
					}
				}
				finally
				{
					if (enumerator2 != null)
					{
						while (true)
						{
							switch (3)
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
			else
			{
				if (!(arrangement is PyramidArrangement))
				{
					throw new NotImplementedException();
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
				Arrangements[0].IsChecked = true;
			}
			arrangement = null;
		}
		ObservableCollection<Arrangement> arrangements = Arrangements;
		Func<Arrangement, bool> predicate;
		if (_Closure_0024__.A == null)
		{
			predicate = (_Closure_0024__.A = [SpecialName] (Arrangement A) => A.IsChecked);
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
			predicate = _Closure_0024__.A;
		}
		if (arrangements.Where(predicate).Count() != 0)
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
			K();
			return;
		}
	}

	private void B(int A)
	{
		IEnumerator<Arrangement> enumerator = default(IEnumerator<Arrangement>);
		try
		{
			ObservableCollection<Arrangement> arrangements = Arrangements;
			Func<Arrangement, bool> predicate;
			if (_Closure_0024__.B == null)
			{
				predicate = (_Closure_0024__.B = [SpecialName] (Arrangement arrangement) => arrangement is RectArrangement);
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
			enumerator = arrangements.Where(predicate).GetEnumerator();
			while (enumerator.MoveNext())
			{
				RectArrangement rectArrangement = (RectArrangement)enumerator.Current;
				if (rectArrangement.Rows.Count == A)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							rectArrangement.IsChecked = true;
							return;
						}
					}
				}
				rectArrangement = null;
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
	}

	private void ReselectShapes(object sender, RoutedEventArgs e)
	{
		try
		{
			ShapeRange.Select();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
	}

	private void BackToParent(object sender, RoutedEventArgs e)
	{
		Action parentCloseAction = ParentCloseAction;
		if (parentCloseAction == null)
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
		parentCloseAction();
	}

	private void ContainerSizeChanged(object sender, SizeChangedEventArgs e)
	{
		if (!e.HeightChanged)
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
			ScrollViewer scrollViewer = scroller;
			Thickness padding;
			if (scroller.ComputedVerticalScrollBarVisibility != Visibility.Visible)
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
				padding = new Thickness(0.0);
			}
			else
			{
				padding = new Thickness(0.0, 0.0, 8.0, 0.0);
			}
			scrollViewer.Padding = padding;
			return;
		}
	}

	private void ArrangeClosed(object sender, RoutedEventArgs e)
	{
		I();
		try
		{
			if (this.m_A != null)
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
					this.m_A.CancelAsync();
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
		this.m_A = null;
		try
		{
			IEnumerator<ShapeItem> enumerator = default(IEnumerator<ShapeItem>);
			try
			{
				enumerator = ReorderShapes.GetEnumerator();
				while (enumerator.MoveNext())
				{
					enumerator.Current.C();
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_006a;
					}
					continue;
					end_IL_006a:
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
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		Containers = null;
		Arrangements = null;
		Container = null;
		ParentCloseAction = null;
		ShapeRange = null;
		ShapeItems = null;
		ReorderShapes = null;
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
			Uri resourceLocator = new Uri(AH.A(69430), UriKind.Relative);
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
		//IL_01c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_01cf: Expected O, but got Unknown
		//IL_0235: Unknown result type (might be due to invalid IL or missing references)
		//IL_023f: Expected O, but got Unknown
		//IL_0263: Unknown result type (might be due to invalid IL or missing references)
		//IL_026d: Expected O, but got Unknown
		//IL_0251: Unknown result type (might be due to invalid IL or missing references)
		//IL_025b: Expected O, but got Unknown
		//IL_0287: Unknown result type (might be due to invalid IL or missing references)
		//IL_0291: Expected O, but got Unknown
		//IL_02a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_02ad: Expected O, but got Unknown
		if (connectionId == 3)
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
					grdLibrary = (Grid)target;
					return;
				}
			}
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
					btnBack = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnSelect = (Button)target;
					return;
				}
			}
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
					scroller = (ScrollViewer)target;
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
					((Grid)target).SizeChanged += ContainerSizeChanged;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkContainer = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					lbxArrange = (ListBox)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					grpReorder = (GroupBox)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			chkReorder = (CheckBox)target;
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
					lbxReorder = (ListBox)target;
					return;
				}
			}
		}
		if (connectionId == 15)
		{
			btnMoveUp = (Button)target;
			btnMoveUp.Click += MoveShapeUp;
			return;
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnMoveDown = (Button)target;
					btnMoveDown.Click += MoveShapeDown;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					grpPrefs = (GroupBox)target;
					return;
				}
			}
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkPreferences = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					numMaxShapes = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 20)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkScale = (CheckBox)target;
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
					cbxStretch = (ComboBox)target;
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
					chkCenter = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 23)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					numColumnSpacing = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 24)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					numRowSpacing = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 25)
		{
			numPadding = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 26)
		{
			cbxCircleAlign = (ComboBox)target;
			return;
		}
		if (connectionId == 27)
		{
			numCircleScale = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 28)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					numRotationAngle = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 29)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkRotate = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 30)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkBestFit = (CheckBox)target;
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
		if (connectionId == 1)
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
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = ToggleButton.CheckedEvent;
			eventSetter.Handler = new RoutedEventHandler(ContainerChecked);
			((Style)target).Setters.Add(eventSetter);
		}
		if (connectionId == 2)
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
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = ToggleButton.CheckedEvent;
			eventSetter.Handler = new RoutedEventHandler(ArrangementChecked);
			((Style)target).Setters.Add(eventSetter);
		}
		if (connectionId == 13)
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
			((RadioButton)target).Checked += ReorderShapeChecked;
		}
		if (connectionId != 14)
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
			((Button)target).Click += DeleteShape;
			return;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
