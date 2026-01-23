using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media.Imaging;
using A;
using MacabacusMacros;
using MacabacusMacros.Libraries;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Agenda;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Pagination;

[DesignerGenerated]
public sealed class wpfPaginate : System.Windows.Controls.UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private readonly int m_A;

	private Microsoft.Office.Interop.PowerPoint.Presentation m_A;

	private BackgroundWorker m_A;

	private FlySheetStyle m_A;

	[CompilerGenerated]
	private List<BaseSlideItem> m_A;

	private ObservableCollection<SlidePair> m_A;

	private int m_B;

	private int m_C;

	private ObservableCollection<int> m_A;

	private int m_D;

	[CompilerGenerated]
	private int m_E;

	[CompilerGenerated]
	private Dictionary<int, BitmapImage> m_A;

	[CompilerGenerated]
	private float m_A;

	[CompilerGenerated]
	private float m_B;

	[CompilerGenerated]
	private List<int> m_A;

	[AccessedThroughProperty("radDuplex")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("radSimplex")]
	private System.Windows.Controls.RadioButton m_B;

	[AccessedThroughProperty("chkFlysheets")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_A;

	[AccessedThroughProperty("lbxBindings")]
	[CompilerGenerated]
	private System.Windows.Controls.ListBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lvPages")]
	private System.Windows.Controls.ListView m_A;

	[AccessedThroughProperty("radViewBindings")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("radViewPages")]
	private System.Windows.Controls.RadioButton m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("btnRefresh")]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkStartAtOne")]
	private System.Windows.Controls.CheckBox m_B;

	[AccessedThroughProperty("chkSequential")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnFinalize")]
	private System.Windows.Controls.Button m_B;

	private bool m_A;

	private List<BaseSlideItem> Slides
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

	public ObservableCollection<SlidePair> SlidePairs
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(100435));
		}
	}

	public int RowsCount
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(AH.A(100456));
		}
	}

	public int ThumbnailHeight
	{
		get
		{
			return this.m_C;
		}
		set
		{
			this.m_C = value;
			A(AH.A(69008));
			K();
		}
	}

	public ObservableCollection<int> Spirals
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(100475));
		}
	}

	public int FacingSlidesCount
	{
		get
		{
			return this.m_D;
		}
		set
		{
			this.m_D = value;
			A(AH.A(100490));
		}
	}

	private int BlankSlidesCount
	{
		[CompilerGenerated]
		get
		{
			return this.m_E;
		}
		[CompilerGenerated]
		set
		{
			this.m_E = value;
		}
	}

	private Dictionary<int, BitmapImage> SlideThumbCache
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

	private float OriginalSlideHeight
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

	private float OriginalSlideWidth
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

	private List<int> OriginalSlideIds
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

	internal virtual System.Windows.Controls.RadioButton radDuplex
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

	internal virtual System.Windows.Controls.RadioButton radSimplex
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

	internal virtual System.Windows.Controls.CheckBox chkFlysheets
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

	internal virtual System.Windows.Controls.ListBox lbxBindings
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

	internal virtual System.Windows.Controls.ListView lvPages
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

	internal virtual System.Windows.Controls.RadioButton radViewBindings
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

	internal virtual System.Windows.Controls.RadioButton radViewPages
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
			RoutedEventHandler value2 = RefreshSlidesClicked;
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

	internal virtual System.Windows.Controls.CheckBox chkStartAtOne
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

	internal virtual System.Windows.Controls.CheckBox chkSequential
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

	internal virtual System.Windows.Controls.Button btnFinalize
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
			RoutedEventHandler value2 = DoFinalize;
			System.Windows.Controls.Button button = this.m_B;
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
	}

	public wpfPaginate()
	{
		base.Unloaded += wpfPaginatePane_Unloaded;
		this.m_A = 200;
		InitializeComponent();
	}

	private void A(string A)
	{
		this.m_A?.Invoke(this, new PropertyChangedEventArgs(A));
	}

	private void wpfPaginatePane_Unloaded(object sender, RoutedEventArgs e)
	{
		A();
	}

	public void ShowPane()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		this.m_A = application.ActivePresentation;
		this.m_A = Behavior.GetPresentationFlysheetStyle(this.m_A);
		chkSequential.IsChecked = KG.A.SequentialSlideNumbers;
		chkStartAtOne.IsChecked = KG.A.SlideNumbersStartAtOne;
		radDuplex.IsChecked = PB.Settings.PaginateDuplex;
		System.Windows.Controls.RadioButton radioButton = radSimplex;
		bool? isChecked = radDuplex.IsChecked;
		bool? isChecked2;
		if (!isChecked.HasValue)
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
			isChecked2 = isChecked;
		}
		else
		{
			isChecked2 = isChecked != true;
		}
		radioButton.IsChecked = isChecked2;
		chkFlysheets.IsChecked = PB.Settings.PaginateFlysheetsFront;
		radViewBindings.IsChecked = PB.Settings.PaginateBindingsView;
		System.Windows.Controls.RadioButton radioButton2 = radViewPages;
		isChecked = radViewBindings.IsChecked;
		radioButton2.IsChecked = (!isChecked) ?? isChecked;
		if (radViewBindings.IsChecked == true)
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
			lbxBindings.Visibility = Visibility.Visible;
			lvPages.Visibility = Visibility.Collapsed;
		}
		else
		{
			lbxBindings.Visibility = Visibility.Collapsed;
			lvPages.Visibility = Visibility.Visible;
		}
		btnFinalize.IsEnabled = true;
		A(application);
		SlideThumbCache = new Dictionary<int, BitmapImage>();
		B();
		application = null;
	}

	public void HidePane()
	{
		if (this.m_A != null)
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
			if (this.m_A.IsBusy)
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
				this.m_A.CancelAsync();
			}
		}
		A();
		this.m_A = null;
		this.m_A = null;
		Slides = null;
		SlidePairs = null;
		OriginalSlideIds = null;
		SlideThumbCache = null;
		Spirals = null;
		GC.Collect();
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Application A)
	{
		B(A);
		lvPages.IsVisibleChanged += ResizeGridViewColumns;
		radDuplex.Checked += DuplexChecked;
		radSimplex.Checked += SimplexChecked;
		radViewBindings.Checked += ViewChanged;
		radViewPages.Checked += ViewChanged;
		chkFlysheets.Checked += FlysheetsCheckedChanged;
		chkFlysheets.Unchecked += FlysheetsCheckedChanged;
	}

	private void A()
	{
		L();
		lvPages.IsVisibleChanged -= ResizeGridViewColumns;
		radDuplex.Checked -= DuplexChecked;
		radSimplex.Checked -= SimplexChecked;
		radViewBindings.Checked -= ViewChanged;
		radViewPages.Checked -= ViewChanged;
		chkFlysheets.Checked -= FlysheetsCheckedChanged;
		chkFlysheets.Unchecked -= FlysheetsCheckedChanged;
	}

	private void ViewChanged(object sender, RoutedEventArgs e)
	{
		if (radViewBindings.IsChecked == true)
		{
			lbxBindings.Visibility = Visibility.Visible;
			lvPages.Visibility = Visibility.Collapsed;
		}
		else
		{
			lvPages.Visibility = Visibility.Visible;
			lbxBindings.Visibility = Visibility.Collapsed;
		}
		B();
		PB.Settings.PaginateBindingsView = radViewBindings.IsChecked.Value;
	}

	private void DuplexChecked(object sender, RoutedEventArgs e)
	{
		chkFlysheets.IsEnabled = true;
		B();
		PB.Settings.PaginateDuplex = true;
	}

	private void SimplexChecked(object sender, RoutedEventArgs e)
	{
		System.Windows.Controls.CheckBox checkBox = chkFlysheets;
		checkBox.IsEnabled = false;
		checkBox.Checked += FlysheetsCheckedChanged;
		checkBox.Unchecked += FlysheetsCheckedChanged;
		checkBox.IsChecked = true;
		checkBox.Checked -= FlysheetsCheckedChanged;
		checkBox.Unchecked -= FlysheetsCheckedChanged;
		_ = null;
		B();
		PB.Settings.PaginateDuplex = false;
	}

	private void FlysheetsCheckedChanged(object sender, RoutedEventArgs e)
	{
		B();
		PB.Settings.PaginateFlysheetsFront = chkFlysheets.IsChecked.Value;
	}

	private void ResizeGridViewColumns(object sender, DependencyPropertyChangedEventArgs e)
	{
		if (!lvPages.IsVisible)
		{
			return;
		}
		IEnumerator<GridViewColumn> enumerator = default(IEnumerator<GridViewColumn>);
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
				enumerator = ((GridView)lvPages.View).Columns.GetEnumerator();
				while (enumerator.MoveNext())
				{
					GridViewColumn current = enumerator.Current;
					if (double.IsNaN(current.Width))
					{
						current.Width = current.ActualWidth;
					}
					current.Width = double.NaN;
				}
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
		}
	}

	private void B()
	{
		Slides = new List<BaseSlideItem>();
		SlidePairs = new ObservableCollection<SlidePair>();
		FacingSlidesCount = 0;
		BlankSlidesCount = 0;
		Master slideMaster = this.m_A.SlideMaster;
		OriginalSlideHeight = slideMaster.Height;
		OriginalSlideWidth = slideMaster.Width;
		checked
		{
			ThumbnailHeight = (int)Math.Round((float)this.m_A * slideMaster.Height / slideMaster.Width);
			slideMaster = null;
			OriginalSlideIds = new List<int>();
			int count = this.m_A.Slides.Count;
			for (int i = 1; i <= count; i++)
			{
				OriginalSlideIds.Add(this.m_A.Slides[i].SlideID);
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
				if (radDuplex.IsChecked == true)
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
					E();
				}
				else
				{
					F();
				}
				if (radViewBindings.IsChecked == true)
				{
					C();
				}
				else
				{
					D();
				}
				J();
				return;
			}
		}
	}

	private void C()
	{
		checked
		{
			int num = Slides.Count - 1;
			int num2 = num;
			for (int i = 0; i <= num2; i++)
			{
				if (i == 0)
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
					G(new EmptySlideItem(), Slides[i]);
				}
				else if (i < num)
				{
					if (unchecked(checked(i + 1) % 2) == 0)
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
						if (Slides[i + 1] is FlysheetSlideItem)
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
							FlysheetSlideItem obj = (FlysheetSlideItem)Slides[i + 1];
							obj.AdornerAlignment = System.Windows.HorizontalAlignment.Right;
							obj.AdornerPadding = new Thickness(3.0, 0.0, 2.0, 0.0);
							_ = null;
						}
						else if (Slides[i + 1] is FacingSlideItem)
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
							FacingSlideItem obj2 = (FacingSlideItem)Slides[i + 1];
							obj2.AdornerAlignment = System.Windows.HorizontalAlignment.Right;
							obj2.AdornerPadding = new Thickness(3.0, 0.0, 2.0, 0.0);
							_ = null;
						}
					}
					G(Slides[i], Slides[i + 1]);
					i++;
				}
				else
				{
					G(Slides[i], new EmptySlideItem());
				}
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
	}

	private void D()
	{
		checked
		{
			int num = Slides.Count - 1;
			int num2 = num;
			for (int i = 0; i <= num2; i++)
			{
				if (i < num)
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
					H(Slides[i], Slides[i + 1]);
					i++;
				}
				else
				{
					H(Slides[i], null);
				}
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
	}

	private void E()
	{
		int count = this.m_A.Slides.Count;
		int num = count;
		checked
		{
			Slide slide;
			for (int i = 1; i <= num; slide = null, i++)
			{
				slide = this.m_A.Slides[i];
				Slide slide2;
				if (!FacingSlides.IsFacingSlide(slide))
				{
					if (A(slide))
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
						Slides.Add(A(slide));
					}
					else
					{
						Slides.Add(A(slide));
					}
					if (i >= count)
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
					slide2 = this.m_A.Slides[i + 1];
					if (!FacingSlides.IsFacingSlide(slide2))
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
						bool flag = A(slide2);
						bool? isChecked = chkFlysheets.IsChecked;
						if (((!isChecked) ?? isChecked) != true)
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
							if (!flag)
							{
								Slides.Add(A());
								goto IL_01ab;
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
						if (flag)
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
							Slides.Add(A(slide2));
						}
						else
						{
							Slides.Add(A(slide2));
						}
						i++;
					}
					else
					{
						Slides.Add(A(slide2));
						i++;
					}
					goto IL_01ab;
				}
				Slides.Add(A());
				Slides.Add(A(slide));
				continue;
				IL_01ab:
				slide2 = null;
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
	}

	private void F()
	{
		int count = this.m_A.Slides.Count;
		int num = count;
		checked
		{
			Slide slide;
			for (int i = 1; i <= num; slide = null, i++)
			{
				slide = this.m_A.Slides[i];
				if (!FacingSlides.IsFacingSlide(slide))
				{
					if (A(slide))
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
						Slides.Add(A(slide));
					}
					else
					{
						Slides.Add(A(slide));
					}
					if (i >= count)
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
					Slide slide2 = this.m_A.Slides[i + 1];
					if (!FacingSlides.IsFacingSlide(slide2))
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
						Slides.Add(A());
					}
					else
					{
						Slides.Add(A(slide2));
						i++;
					}
					slide2 = null;
					continue;
				}
				if (i != 1)
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
					if (!FacingSlides.IsFacingSlide(this.m_A.Slides[i - 1]))
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
				}
				Slides.Add(A());
				Slides.Add(A(slide));
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
	}

	private void G(BaseSlideItem A, BaseSlideItem B)
	{
		SlidePairs.Add(new SlidePair(A, B));
	}

	private void H(BaseSlideItem A, BaseSlideItem B)
	{
		SlidePairs.Add(new SlidePair(A, B, checked(SlidePairs.Count + 1)));
	}

	private RegularSlideItem A(Slide A)
	{
		return new RegularSlideItem(A, this.A(A));
	}

	private FlysheetSlideItem A(Slide A)
	{
		return new FlysheetSlideItem(A, this.A(A));
	}

	private FacingSlideItem A(Slide A)
	{
		checked
		{
			FacingSlidesCount++;
			return new FacingSlideItem(A, this.A(A));
		}
	}

	private BlankSlideItem A()
	{
		checked
		{
			BlankSlidesCount++;
			return new BlankSlideItem();
		}
	}

	private BitmapImage A(Slide A)
	{
		BitmapImage value = null;
		SlideThumbCache.TryGetValue(A.SlideIndex, out value);
		return value;
	}

	private bool A(Slide A)
	{
		return !BlankSlides.IsSectionFlysheet(A, this.m_A, this.m_A);
	}

	private void RefreshSlidesClicked(object sender, RoutedEventArgs e)
	{
		I();
	}

	private void I()
	{
		SlideThumbCache.Clear();
		B();
	}

	private void J()
	{
		this.m_A = new BackgroundWorker();
		BackgroundWorker a = this.m_A;
		a.WorkerReportsProgress = false;
		a.WorkerSupportsCancellation = true;
		a.DoWork += AsynchLoadThumbnails;
		a.RunWorkerCompleted += AsynchLoadComplete;
		_ = null;
		this.m_A.RunWorkerAsync();
	}

	private void AsynchLoadThumbnails(object sender, DoWorkEventArgs e)
	{
		Random b = new Random();
		using (List<BaseSlideItem>.Enumerator enumerator = Slides.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				BaseSlideItem current = enumerator.Current;
				if (current is RegularSlideItem)
				{
					RegularSlideItem regularSlideItem = (RegularSlideItem)current;
					regularSlideItem.Thumbnail = A(regularSlideItem.Slide, b);
					regularSlideItem = null;
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
		b = null;
	}

	private void AsynchLoadComplete(object sender, RunWorkerCompletedEventArgs e)
	{
	}

	private BitmapImage A(Slide A, Random B)
	{
		int slideIndex = A.SlideIndex;
		if (!SlideThumbCache.TryGetValue(slideIndex, out var value))
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
				string text = Path.Combine(Path.GetTempPath(), Base.RandomStringGenerator(B) + AH.A(100525));
				int scaleHeight = clsPowerPoint.ScaleHeight(600, this.m_A);
				A.Export(text, AH.A(63328), 600, scaleHeight);
				value = new BitmapImage();
				BitmapImage bitmapImage = value;
				bitmapImage.BeginInit();
				bitmapImage.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
				bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
				bitmapImage.DecodePixelWidth = this.m_A;
				bitmapImage.DecodePixelHeight = ThumbnailHeight;
				bitmapImage.UriSource = new Uri(text);
				bitmapImage.EndInit();
				bitmapImage.Freeze();
				_ = null;
				SlideThumbCache.Add(slideIndex, value);
				File.Delete(text);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		return value;
	}

	private void K()
	{
		checked
		{
			double a = Math.Floor((double)(ThumbnailHeight - 12) / 9.0) - 2.0;
			Spirals = new ObservableCollection<int>();
			int num = (int)Math.Round(a);
			for (int i = 1; i <= num; i++)
			{
				Spirals.Add(i);
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
				return;
			}
		}
	}

	private void DoFinalize(object sender, RoutedEventArgs e)
	{
		CustomLayout customLayout = null;
		int num = 0;
		if (this.m_A != null)
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
			if (this.m_A.IsBusy)
			{
				this.m_A.CancelAsync();
			}
		}
		if (this.m_A.Final)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					Pane.MarkedFinalWarning();
					return;
				}
			}
		}
		if (A())
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					B(AH.A(100534));
					return;
				}
			}
		}
		if (Forms.OkCancelMessage2(AH.A(100692)) != DialogResult.OK)
		{
			return;
		}
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			IEnumerator enumerator2 = default(IEnumerator);
			IEnumerator enumerator3 = default(IEnumerator);
			IEnumerator enumerator4 = default(IEnumerator);
			IEnumerator enumerator5 = default(IEnumerator);
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				if (this.m_A.Saved == MsoTriState.msoFalse)
				{
					if (Forms.OkCancelMessage2(this.m_A.Name + AH.A(100846)) != DialogResult.OK)
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
						break;
					}
					this.m_A.Save();
				}
				this.m_A.SaveAs(Path.GetDirectoryName(this.m_A.FullName) + Conversions.ToString(Path.DirectorySeparatorChar) + Path.GetFileNameWithoutExtension(this.m_A.Name) + AH.A(100965));
				this.m_A.Application.StartNewUndoEntry();
				if (chkSequential.IsChecked == true)
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
					{
						enumerator = this.m_A.Slides.GetEnumerator();
						try
						{
							while (enumerator.MoveNext())
							{
								Slide slide = (Slide)enumerator.Current;
								if (Helpers.GetSlideType(slide) != SlideType.Blank)
								{
									try
									{
										enumerator2 = slide.Shapes.GetEnumerator();
										while (enumerator2.MoveNext())
										{
											Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
											try
											{
												if (!Numbers.IsSlideNumberPlaceholder(shape))
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
													if (num == 0)
													{
														int num2;
														if (chkStartAtOne.IsChecked != true)
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
															num2 = slide.SlideIndex;
														}
														else
														{
															num2 = 1;
														}
														num = num2;
													}
													else
													{
														num++;
													}
													shape.TextFrame.TextRange.Text = num.ToString();
													break;
												}
												break;
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
										if (enumerator2 is IDisposable)
										{
											while (true)
											{
												switch (4)
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
								FacingSlides.A(slide);
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_0295;
								}
								continue;
								end_IL_0295:
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
				}
				else if (chkStartAtOne.IsChecked != true)
				{
					{
						enumerator3 = this.m_A.Slides.GetEnumerator();
						try
						{
							while (enumerator3.MoveNext())
							{
								Slide obj = (Slide)enumerator3.Current;
								SlideNumbers.Freeze(obj);
								FacingSlides.A(obj);
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_0456;
								}
								continue;
								end_IL_0456:
								break;
							}
						}
						finally
						{
							IDisposable disposable2 = enumerator3 as IDisposable;
							if (disposable2 != null)
							{
								disposable2.Dispose();
							}
						}
					}
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
					try
					{
						enumerator4 = this.m_A.Slides.GetEnumerator();
						while (enumerator4.MoveNext())
						{
							Slide slide2 = (Slide)enumerator4.Current;
							if (Helpers.GetSlideType(slide2) != SlideType.Blank)
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
								if (num > 0)
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
									num++;
								}
								try
								{
									enumerator5 = slide2.Shapes.GetEnumerator();
									while (true)
									{
										if (enumerator5.MoveNext())
										{
											Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator5.Current;
											try
											{
												if (!Numbers.IsSlideNumberPlaceholder(shape2))
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
													if (num == 0)
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
														num++;
													}
													shape2.TextFrame.TextRange.Text = num.ToString();
													break;
												}
												break;
											}
											catch (Exception ex3)
											{
												ProjectData.SetProjectError(ex3);
												Exception ex4 = ex3;
												ProjectData.ClearProjectError();
											}
											continue;
										}
										while (true)
										{
											switch (3)
											{
											case 0:
												break;
											default:
												goto end_IL_03b3;
											}
											continue;
											end_IL_03b3:
											break;
										}
										break;
									}
								}
								finally
								{
									if (enumerator5 is IDisposable)
									{
										while (true)
										{
											switch (1)
											{
											case 0:
												continue;
											}
											(enumerator5 as IDisposable).Dispose();
											break;
										}
									}
								}
							}
							FacingSlides.A(slide2);
						}
					}
					finally
					{
						if (enumerator4 is IDisposable)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								(enumerator4 as IDisposable).Dispose();
								break;
							}
						}
					}
				}
				KG.A.SequentialSlideNumbers = chkSequential.IsChecked.Value;
				KG.A.SlideNumbersStartAtOne = chkStartAtOne.IsChecked.Value;
				if (BlankSlidesCount > 0)
				{
					if (radDuplex.IsChecked == true)
					{
						customLayout = BlankSlides.GetBlankLayout(this.m_A);
					}
					if (customLayout == null)
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
						customLayout = BlankSlides.CreateBlankLayout(this.m_A);
					}
					A(customLayout);
				}
				this.m_A.Final = true;
				btnFinalize.IsEnabled = false;
				clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)4, AH.A(100978));
				customLayout = null;
				return;
			}
		}
	}

	private void A(CustomLayout A)
	{
		checked
		{
			int num = Slides.Count - 1;
			for (int i = 0; i <= num; i++)
			{
				if (Slides[i] is BlankSlideItem)
				{
					this.m_A.Slides.AddSlide(((RegularSlideItem)Slides[i + 1]).Slide.SlideIndex, A);
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
				return;
			}
		}
	}

	private void A(int A)
	{
		this.m_A.Slides.Add(A, PpSlideLayout.ppLayoutBlank).Shapes.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).Delete();
		_ = null;
	}

	private void SlideMouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
	{
		A((Border)sender, Visibility.Visible);
	}

	private void SlideMouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
	{
		A((Border)sender, Visibility.Hidden);
	}

	private void A(Border A, Visibility B)
	{
		try
		{
			((RegularSlideItem)A.DataContext).ToggleVisibility = B;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void MarkSlideClick(object sender, RoutedEventArgs e)
	{
		RegularSlideItem obj = (RegularSlideItem)((System.Windows.Controls.Button)sender).DataContext;
		obj.Slide.Application.StartNewUndoEntry();
		FacingSlides.A(obj.Slide, this.m_A);
		B();
	}

	private void UnmarkSlideClick(object sender, RoutedEventArgs e)
	{
		FacingSlideItem obj = (FacingSlideItem)((System.Windows.Controls.Button)sender).DataContext;
		obj.Slide.Application.StartNewUndoEntry();
		FacingSlides.A(obj.Slide);
		B();
	}

	private void B(Microsoft.Office.Interop.PowerPoint.Application A)
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56735)).AddEventHandler(A, new EApplication_SlideSelectionChangedEventHandler(this.A));
	}

	private void L()
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56735)).RemoveEventHandler(NG.A.Application, new EApplication_SlideSelectionChangedEventHandler(A));
	}

	private void A(SlideRange A)
	{
		if (!this.A())
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
			I();
			return;
		}
	}

	private bool A()
	{
		int count = this.m_A.Slides.Count;
		if (count != OriginalSlideIds.Count)
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
					return true;
				}
			}
		}
		int num = count;
		checked
		{
			for (int i = 1; i <= num; i++)
			{
				if (this.m_A.Slides[i].SlideID == OriginalSlideIds[i - 1])
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
					return true;
				}
			}
			Master slideMaster = this.m_A.SlideMaster;
			if (slideMaster.Height == OriginalSlideHeight)
			{
				if (slideMaster.Width == OriginalSlideWidth)
				{
					slideMaster = null;
					return false;
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
			return true;
		}
	}

	private void B(string A)
	{
		Forms.WarningMessage(A);
	}

	private void C(string A)
	{
		Forms.ErrorMessage(A);
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
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_A = true;
			Uri resourceLocator = new Uri(AH.A(101017), UriKind.Relative);
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
	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 4)
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
					radDuplex = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			radSimplex = (System.Windows.Controls.RadioButton)target;
			return;
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkFlysheets = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					lbxBindings = (System.Windows.Controls.ListBox)target;
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
					lvPages = (System.Windows.Controls.ListView)target;
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
					radViewBindings = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
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
					radViewPages = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			btnRefresh = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkStartAtOne = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 13:
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				chkSequential = (System.Windows.Controls.CheckBox)target;
				return;
			}
		case 14:
			btnFinalize = (System.Windows.Controls.Button)target;
			break;
		default:
			this.m_A = true;
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
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = UIElement.MouseEnterEvent;
			eventSetter.Handler = new System.Windows.Input.MouseEventHandler(SlideMouseEnter);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.MouseLeaveEvent;
			eventSetter.Handler = new System.Windows.Input.MouseEventHandler(SlideMouseLeave);
			((Style)target).Setters.Add(eventSetter);
		}
		if (connectionId == 2)
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
			((System.Windows.Controls.Button)target).Click += MarkSlideClick;
		}
		if (connectionId == 3)
		{
			((System.Windows.Controls.Button)target).Click += UnmarkSlideClick;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
