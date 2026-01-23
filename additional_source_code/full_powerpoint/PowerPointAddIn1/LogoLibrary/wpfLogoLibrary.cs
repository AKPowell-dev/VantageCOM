using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Markup;
using System.Windows.Media;
using A;
using MacabacusMacros;
using MacabacusMacros.LogoLibrary.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes.Arrange;
using SharpVectors.Converters;
using SharpVectors.Runtime;

namespace PowerPointAddIn1.LogoLibrary;

[DesignerGenerated]
public sealed class wpfLogoLibrary : UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	public enum Theme
	{
		Auto,
		Light,
		Dark
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<LogoItem, bool> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(LogoItem A)
		{
			return A.IsChecked;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	[CompilerGenerated]
	private AutoSuggest m_A;

	private ListCollectionView m_A;

	[CompilerGenerated]
	private ObservableCollection<LogoItem> m_A;

	private UserControl m_A;

	private UserControl m_B;

	private int m_A;

	private string m_A;

	private SolidColorBrush m_A;

	private SolidColorBrush m_B;

	[CompilerGenerated]
	private Theme m_A;

	private ListSortDirection? m_A;

	private Geometry m_A;

	[AccessedThroughProperty("header")]
	[CompilerGenerated]
	private DockPanel m_A;

	[AccessedThroughProperty("txtSearch")]
	[CompilerGenerated]
	private TextBox m_A;

	[AccessedThroughProperty("popSuggest")]
	[CompilerGenerated]
	private Popup m_A;

	[AccessedThroughProperty("spinner")]
	[CompilerGenerated]
	private Grid m_A;

	[AccessedThroughProperty("btnInsert")]
	[CompilerGenerated]
	private Button m_A;

	private bool m_A;

	public AutoSuggest AutoSuggest
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

	public ListCollectionView SourceCollection
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(10961));
		}
	}

	private ObservableCollection<LogoItem> Logos
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

	public UserControl ErrorView
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(10162));
		}
	}

	public UserControl ArrangeView
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(AH.A(10994));
		}
	}

	private int SelectedLogoCount
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			string insertButtonText;
			if (value != 1)
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
				insertButtonText = AH.A(11017) + value + AH.A(11032);
			}
			else
			{
				insertButtonText = AH.A(11017) + value + AH.A(11045);
			}
			InsertButtonText = insertButtonText;
		}
	}

	public string InsertButtonText
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(11056));
		}
	}

	public SolidColorBrush ThemeLeftHalf
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(11089));
		}
	}

	public SolidColorBrush ThemeRightHalf
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(AH.A(11116));
		}
	}

	private Theme CurrentTheme
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

	private ListSortDirection? CurrentSort
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			if (!value.HasValue)
			{
				return;
			}
			if (value.Value == ListSortDirection.Ascending)
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
						AlphaSortIcon = Geometry.Parse(AH.A(11145));
						return;
					}
				}
			}
			if (value.Value != ListSortDirection.Descending)
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
				AlphaSortIcon = Geometry.Parse(AH.A(11639));
				return;
			}
		}
	}

	public Geometry AlphaSortIcon
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(12141));
		}
	}

	internal virtual DockPanel header
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
			this.m_A = value;
		}
	}

	internal virtual Popup popSuggest
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

	internal virtual Grid spinner
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

	internal virtual Button btnInsert
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
			RoutedEventHandler value2 = InsertLogo;
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
	}

	public wpfLogoLibrary()
	{
		base.Loaded += wpfLogoLibrary_Loaded;
		base.Unloaded += wpfLogoLibrary_Unloaded;
		this.m_A = null;
		this.m_A = null;
		this.m_B = null;
		this.m_A = null;
		this.m_B = null;
		this.m_A = null;
		InitializeComponent();
		bool flag = false;
		DockPanel dockPanel = header;
		int visibility;
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
		dockPanel.Visibility = (Visibility)visibility;
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

	private void wpfLogoLibrary_Loaded(object sender, RoutedEventArgs e)
	{
	}

	private void wpfLogoLibrary_Unloaded(object sender, RoutedEventArgs e)
	{
		A();
	}

	private void A()
	{
		AutoSuggest autoSuggest = AutoSuggest;
		if (autoSuggest == null)
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
			autoSuggest.ShutDown();
		}
		SourceCollection = null;
		Logos = null;
		ErrorView = null;
		ArrangeView = null;
		AutoSuggest = null;
		ThemeLeftHalf = null;
		ThemeRightHalf = null;
		AlphaSortIcon = null;
	}

	public void ShowPane()
	{
		//IL_0043: Unknown result type (might be due to invalid IL or missing references)
		//IL_004d: Expected O, but got Unknown
		Logos = new ObservableCollection<LogoItem>();
		AutoSuggest = new AutoSuggest(txtSearch, popSuggest, spinner, base.Dispatcher, (Action<string>)B, NG.A.Application, false);
		AutoSuggest.ClearAction = H;
		AutoSuggest.AddAction = [SpecialName] (object A) =>
		{
			//IL_0046: Unknown result type (might be due to invalid IL or missing references)
			//IL_0050: Expected O, but got Unknown
			//IL_0033: Unknown result type (might be due to invalid IL or missing references)
			//IL_001d: Unknown result type (might be due to invalid IL or missing references)
			if (CurrentTheme == Theme.Light)
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
				((LogoItem)A).SetThemeLight();
			}
			else if (CurrentTheme == Theme.Dark)
			{
				((LogoItem)A).SetThemeDark();
			}
			Logos.Add((LogoItem)A);
		};
		btnInsert.IsEnabled = false;
		B();
		SelectedLogoCount = 0;
		CurrentSort = null;
		SourceCollection = (ListCollectionView)CollectionViewSource.GetDefaultView(Logos);
	}

	public void HidePane()
	{
		A();
		G();
	}

	private void ThemeToggle(object sender, RoutedEventArgs e)
	{
		Theme currentTheme = CurrentTheme;
		if (currentTheme != Theme.Auto)
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
					if (currentTheme != Theme.Light)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								B();
								return;
							}
						}
					}
					D();
					return;
				}
			}
		}
		C();
	}

	private void B()
	{
		object obj = System.Windows.Media.ColorConverter.ConvertFromString(AH.A(12168));
		System.Windows.Media.Color color;
		if (obj == null)
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
			color = default(System.Windows.Media.Color);
		}
		else
		{
			color = (System.Windows.Media.Color)obj;
		}
		ThemeLeftHalf = new SolidColorBrush(color);
		object obj2 = System.Windows.Media.ColorConverter.ConvertFromString(AH.A(12183));
		System.Windows.Media.Color color2;
		if (obj2 == null)
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
			color2 = default(System.Windows.Media.Color);
		}
		else
		{
			color2 = (System.Windows.Media.Color)obj2;
		}
		ThemeRightHalf = new SolidColorBrush(color2);
		IEnumerator<LogoItem> enumerator = default(IEnumerator<LogoItem>);
		try
		{
			enumerator = Logos.GetEnumerator();
			while (enumerator.MoveNext())
			{
				enumerator.Current.SetThemeAuto();
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_00a4;
				}
				continue;
				end_IL_00a4:
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
		CurrentTheme = Theme.Auto;
	}

	private void C()
	{
		object obj = System.Windows.Media.ColorConverter.ConvertFromString(AH.A(12168));
		ThemeLeftHalf = new SolidColorBrush((obj != null) ? ((System.Windows.Media.Color)obj) : default(System.Windows.Media.Color));
		object obj2 = System.Windows.Media.ColorConverter.ConvertFromString(AH.A(12168));
		System.Windows.Media.Color color;
		if (obj2 == null)
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
			color = default(System.Windows.Media.Color);
		}
		else
		{
			color = (System.Windows.Media.Color)obj2;
		}
		ThemeRightHalf = new SolidColorBrush(color);
		IEnumerator<LogoItem> enumerator = default(IEnumerator<LogoItem>);
		try
		{
			enumerator = Logos.GetEnumerator();
			while (enumerator.MoveNext())
			{
				enumerator.Current.SetThemeLight();
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
		CurrentTheme = Theme.Light;
	}

	private void D()
	{
		object obj = System.Windows.Media.ColorConverter.ConvertFromString(AH.A(12183));
		System.Windows.Media.Color color;
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
			color = default(System.Windows.Media.Color);
		}
		else
		{
			color = (System.Windows.Media.Color)obj;
		}
		ThemeLeftHalf = new SolidColorBrush(color);
		object obj2 = System.Windows.Media.ColorConverter.ConvertFromString(AH.A(12183));
		System.Windows.Media.Color color2;
		if (obj2 == null)
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
			color2 = default(System.Windows.Media.Color);
		}
		else
		{
			color2 = (System.Windows.Media.Color)obj2;
		}
		ThemeRightHalf = new SolidColorBrush(color2);
		IEnumerator<LogoItem> enumerator = default(IEnumerator<LogoItem>);
		try
		{
			enumerator = Logos.GetEnumerator();
			while (enumerator.MoveNext())
			{
				enumerator.Current.SetThemeDark();
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
		CurrentTheme = Theme.Dark;
	}

	private void LogosChanged(object sender, NotifyCollectionChangedEventArgs e)
	{
		if (SourceCollection == null)
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
			throw new NotImplementedException();
		}
	}

	private void AlphaToggle(object sender, RoutedEventArgs e)
	{
		if (Logos.Count <= 0)
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
			if (SelectedLogoCount > 0)
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
				G();
			}
			if (CurrentSort.HasValue)
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
				if (CurrentSort.Value != ListSortDirection.Descending)
				{
					CurrentSort = ListSortDirection.Descending;
					goto IL_008c;
				}
			}
			CurrentSort = ListSortDirection.Ascending;
			goto IL_008c;
			IL_008c:
			_ = Logos[0];
			SourceCollection.SortDescriptions.Clear();
			SourceCollection.SortDescriptions.Add(new SortDescription(AH.A(12198), ListSortDirection.Descending));
			SourceCollection.SortDescriptions.Add(new SortDescription(AH.A(12217), CurrentSort.Value));
			return;
		}
	}

	private void SizeToggle(object sender, RoutedEventArgs e)
	{
		throw new NotImplementedException();
	}

	private bool A(object A)
	{
		return true;
	}

	private void LogoChecked(object sender, RoutedEventArgs e)
	{
		E();
		btnInsert.IsEnabled = true;
	}

	private void LogoUnchecked(object sender, RoutedEventArgs e)
	{
		//IL_0029: Unknown result type (might be due to invalid IL or missing references)
		//IL_002f: Expected O, but got Unknown
		E();
		btnInsert.IsEnabled = SelectedLogoCount > 0;
		LogoItem val = (LogoItem)((CheckBox)sender).DataContext;
		string right;
		try
		{
			right = txtSearch.Tag.ToString();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			right = "";
			ProjectData.ClearProjectError();
		}
		if (Operators.CompareString(val.BrandId, right, TextCompare: false) != 0)
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
			Logos.Remove(val);
		}
		val = null;
	}

	private void E()
	{
		ObservableCollection<LogoItem> logos = Logos;
		Func<LogoItem, bool> predicate;
		if (_Closure_0024__.A == null)
		{
			predicate = (_Closure_0024__.A = [SpecialName] (LogoItem A) => A.IsChecked);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			predicate = _Closure_0024__.A;
		}
		SelectedLogoCount = logos.Where(predicate).Count();
	}

	private void InsertLogo(object sender, RoutedEventArgs e)
	{
		//IL_010e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0115: Expected O, but got Unknown
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Slide slide = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		int num = 0;
		try
		{
			slide = application.ActiveWindow.Selection.SlideRange[1];
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		checked
		{
			if (slide != null)
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
				try
				{
					Selection selection = application.ActiveWindow.Selection;
					if (selection.Type == PpSelectionType.ppSelectionShapes)
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
						if (selection.ShapeRange.Count == 1)
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
							shape = selection.ShapeRange[1];
						}
					}
					selection = null;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				application.StartNewUndoEntry();
				application.ActiveWindow.Selection.Unselect();
				int num2 = SourceCollection.Count - 1;
				int num3 = 0;
				while (true)
				{
					if (num3 <= num2)
					{
						LogoItem val = (LogoItem)SourceCollection.GetItemAt(num3);
						if (val.IsChecked)
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
								slide.Shapes.AddPicture2(val.ImageUrl, MsoTriState.msoFalse, MsoTriState.msoTrue, 0f, 0f).Select(MsoTriState.msoFalse);
								num++;
								clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)6, AH.A(12236) + val.ImageFormat);
							}
							catch (Exception ex5)
							{
								ProjectData.SetProjectError(ex5);
								Exception ex6 = ex5;
								B(ex6.Message);
								ProjectData.ClearProjectError();
								break;
							}
							finally
							{
							}
						}
						num3++;
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
					break;
				}
				if (ErrorView == null && num > 1)
				{
					RectangleF? rect = null;
					if (shape != null)
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
						Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
						rect = new RectangleF(shape2.Left, shape2.Top, shape2.Width, shape2.Height);
						shape2 = null;
					}
					ArrangeView = new wpfArrange(rect, F);
				}
				shape = null;
				slide = null;
			}
			application = null;
		}
	}

	private void F()
	{
		ArrangeView = null;
	}

	private void ClearSearchButtonClick(object sender, RoutedEventArgs e)
	{
		G();
	}

	private void G()
	{
		txtSearch.Clear();
	}

	private void H()
	{
		checked
		{
			for (int i = Logos.Count - 1; i >= 0; i += -1)
			{
				if (!Logos[i].IsChecked)
				{
					Logos.RemoveAt(i);
				}
			}
			txtSearch.Tag = null;
		}
	}

	private void B(string A)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		ErrorView = (UserControl)new wpfError(A, (Action<object, RoutedEventArgs>)DismissError);
	}

	private void DismissError(object sender, RoutedEventArgs e)
	{
		ErrorView = null;
	}

	private void DismissErrorAndClosePane(object sender, RoutedEventArgs e)
	{
		Pane.B();
		ErrorView = null;
	}

	private void CloseView(object sender, RoutedEventArgs e)
	{
		Pane.B();
	}

	private void BadSvgError(object sender, SvgErrorArgs e)
	{
		B(e.Message);
		e.Handled = true;
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
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_A = true;
			Uri resourceLocator = new Uri(AH.A(12279), UriKind.Relative);
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
		if (connectionId == 5)
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
					header = (DockPanel)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			((Button)target).Click += CloseView;
			return;
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
					txtSearch = (TextBox)target;
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
					((Button)target).Click += ThemeToggle;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					((Button)target).Click += AlphaToggle;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					popSuggest = (Popup)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					spinner = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnInsert = (Button)target;
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

	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
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
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = ButtonBase.ClickEvent;
			eventSetter.Handler = new RoutedEventHandler(ClearSearchButtonClick);
			((Style)target).Setters.Add(eventSetter);
		}
		if (connectionId == 2)
		{
			((CheckBox)target).Checked += LogoChecked;
			((CheckBox)target).Unchecked += LogoUnchecked;
		}
		if (connectionId == 3)
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
			((CheckBox)target).Checked += LogoChecked;
			((CheckBox)target).Unchecked += LogoUnchecked;
		}
		if (connectionId != 4)
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
			((SvgViewbox)target).Error += BadSvgError;
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
	private void A(object A)
	{
		//IL_0046: Unknown result type (might be due to invalid IL or missing references)
		//IL_0050: Expected O, but got Unknown
		//IL_0033: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		if (CurrentTheme == Theme.Light)
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
			((LogoItem)A).SetThemeLight();
		}
		else if (CurrentTheme == Theme.Dark)
		{
			((LogoItem)A).SetThemeDark();
		}
		Logos.Add((LogoItem)A);
	}
}
