using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Markup;
using System.Windows.Media;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Fix;
using PowerPointAddIn1.DeckCheck.Reformat;
using PowerPointAddIn1.Shapes;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Colors.Recolor;

[DesignerGenerated]
public sealed class wpfRecolor : System.Windows.Controls.UserControl, IComponentConnector, IStyleConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Tuple<int, IndexedObject>, int> A;

		public static Func<int, int> A;

		public static Func<IGrouping<int, int>, A<int, int>> A;

		public static Func<A<int, int>, int> A;

		public static Func<A<int, int>, int> B;

		public static Func<ColorPair, bool> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal int A(Tuple<int, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal int A(int A)
		{
			return A;
		}

		[SpecialName]
		internal A<int, int> A(IGrouping<int, int> A)
		{
			return new A<int, int>(A.Key, A.Count());
		}

		[SpecialName]
		internal int A(A<int, int> A)
		{
			return A.Count;
		}

		[SpecialName]
		internal int B(A<int, int> A)
		{
			return A.Color;
		}

		[SpecialName]
		internal bool A(ColorPair A)
		{
			if (A.NewColor.HasValue)
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
						return A.NewColor.Value != A.OldColor;
					}
				}
			}
			return false;
		}
	}

	private Microsoft.Office.Interop.PowerPoint.Application m_A;

	private Selection m_A;

	private List<Microsoft.Office.Interop.PowerPoint.Shape> m_A;

	private List<Tuple<int, IndexedObject>> m_A;

	private ObservableCollection<ColorPair> m_A;

	private List<object> m_A;

	private bool m_A;

	[AccessedThroughProperty("radShape")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_A;

	[AccessedThroughProperty("radSlide")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_B;

	[AccessedThroughProperty("radPresentation")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_C;

	[AccessedThroughProperty("chkFont")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFill")]
	private System.Windows.Controls.CheckBox m_B;

	[AccessedThroughProperty("chkBorder")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_C;

	[AccessedThroughProperty("chkLayouts")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_D;

	[AccessedThroughProperty("scroller")]
	[CompilerGenerated]
	private ScrollViewer m_A;

	[AccessedThroughProperty("icColors")]
	[CompilerGenerated]
	private ItemsControl m_A;

	[AccessedThroughProperty("radRgb")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_D;

	[AccessedThroughProperty("radHex")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton E;

	[AccessedThroughProperty("btnRefresh")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[AccessedThroughProperty("btnRecolor")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	private bool m_B;

	internal virtual System.Windows.Controls.RadioButton radShape
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

	internal virtual System.Windows.Controls.RadioButton radSlide
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

	internal virtual System.Windows.Controls.RadioButton radPresentation
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

	internal virtual System.Windows.Controls.CheckBox chkFont
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

	internal virtual System.Windows.Controls.CheckBox chkBorder
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

	internal virtual System.Windows.Controls.CheckBox chkLayouts
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

	internal virtual ItemsControl icColors
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

	internal virtual System.Windows.Controls.RadioButton radRgb
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

	internal virtual System.Windows.Controls.RadioButton radHex
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			E = value;
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
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnRecolor
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
			RoutedEventHandler value2 = btnRecolor_Click;
			System.Windows.Controls.Button button = this.m_B;
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
			this.m_B = value;
			button = this.m_B;
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

	public wpfRecolor()
	{
		base.Unloaded += wpfRecolor_Unloaded;
		this.m_A = false;
		InitializeComponent();
		this.m_A = NG.A.Application;
	}

	private void wpfRecolor_Unloaded(object sender, RoutedEventArgs e)
	{
		B();
		this.m_A = null;
	}

	public void ShowPane()
	{
		Selection selection = null;
		try
		{
			selection = this.m_A.ActiveWindow.Selection;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (selection != null)
		{
			PpSelectionType type = selection.Type;
			if (type != PpSelectionType.ppSelectionSlides)
			{
				if (type == PpSelectionType.ppSelectionShapes)
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
					radShape.IsChecked = true;
				}
				else
				{
					radPresentation.IsChecked = true;
					chkLayouts.IsChecked = true;
				}
			}
			else
			{
				radSlide.IsChecked = true;
			}
			selection = null;
		}
		else
		{
			radPresentation.IsChecked = true;
			chkLayouts.IsChecked = true;
		}
		A();
		C();
	}

	public void HidePane()
	{
		B();
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
	}

	private void A(Selection A)
	{
		this.m_A = A;
		bool? isChecked = radPresentation.IsChecked;
		bool? flag;
		if (!isChecked.HasValue)
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
			flag = isChecked;
		}
		else
		{
			flag = isChecked != true;
		}
		isChecked = flag;
		if (isChecked == true)
		{
			C();
		}
	}

	private void A()
	{
		radShape.Checked += ScopeChanged;
		radShape.Unchecked += ScopeChanged;
		radSlide.Checked += ScopeChanged;
		radSlide.Unchecked += ScopeChanged;
		radPresentation.Checked += ScopeChanged;
		radPresentation.Unchecked += ScopeChanged;
		chkLayouts.Checked += ScopeChanged;
		chkLayouts.Unchecked += ScopeChanged;
		chkFont.Checked += ScopeChanged;
		chkFont.Unchecked += ScopeChanged;
		chkFill.Checked += ScopeChanged;
		chkFill.Unchecked += ScopeChanged;
		chkBorder.Checked += ScopeChanged;
		chkBorder.Unchecked += ScopeChanged;
		radRgb.Checked += ColorFormatChanged;
		radHex.Checked += ColorFormatChanged;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(A));
	}

	private void B()
	{
		radShape.Checked -= ScopeChanged;
		radShape.Unchecked -= ScopeChanged;
		radSlide.Checked -= ScopeChanged;
		radSlide.Unchecked -= ScopeChanged;
		radPresentation.Checked -= ScopeChanged;
		radPresentation.Unchecked -= ScopeChanged;
		chkLayouts.Checked -= ScopeChanged;
		chkLayouts.Unchecked -= ScopeChanged;
		chkFont.Checked -= ScopeChanged;
		chkFont.Unchecked -= ScopeChanged;
		chkFill.Checked -= ScopeChanged;
		chkFill.Unchecked -= ScopeChanged;
		chkBorder.Checked -= ScopeChanged;
		chkBorder.Unchecked -= ScopeChanged;
		radRgb.Checked -= ColorFormatChanged;
		radHex.Checked -= ColorFormatChanged;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(A));
	}

	private void C()
	{
		Selection selection = null;
		try
		{
			selection = this.m_A.ActiveWindow.Selection;
			if (selection == null)
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
						throw new Exception();
					}
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
			return;
		}
		this.m_A = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
		this.m_A = new List<Tuple<int, IndexedObject>>();
		this.m_A = new ObservableCollection<ColorPair>();
		btnRecolor.IsEnabled = false;
		try
		{
			if (radShape.IsChecked == true)
			{
				IEnumerator enumerator = default(IEnumerator);
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					if (selection.Type == PpSelectionType.ppSelectionNone || selection.Type == PpSelectionType.ppSelectionSlides)
					{
						break;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						try
						{
							enumerator = PowerPointAddIn1.Shapes.Base.SelectedShapes(selection).GetEnumerator();
							while (enumerator.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape b = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
								A(selection.SlideRange[1], b);
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
						break;
					}
					break;
				}
			}
			else if (radSlide.IsChecked == true)
			{
				IEnumerator enumerator3 = default(IEnumerator);
				IEnumerator enumerator4 = default(IEnumerator);
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					List<Slide> list = PowerPointAddIn1.Slides.Helpers.A(this.m_A);
					foreach (Slide item in list)
					{
						try
						{
							enumerator3 = item.Shapes.GetEnumerator();
							while (enumerator3.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape b2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator3.Current;
								A(item, b2);
							}
						}
						finally
						{
							if (enumerator3 is IDisposable)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									(enumerator3 as IDisposable).Dispose();
									break;
								}
							}
						}
						if (chkLayouts.IsChecked != true)
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
						try
						{
							enumerator4 = item.CustomLayout.Shapes.GetEnumerator();
							while (enumerator4.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape b3 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator4.Current;
								A(item.CustomLayout, b3);
							}
						}
						finally
						{
							if (enumerator4 is IDisposable)
							{
								while (true)
								{
									switch (2)
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
					list.Clear();
					list = null;
					break;
				}
			}
			else if (radPresentation.IsChecked == true)
			{
				IEnumerator enumerator5 = default(IEnumerator);
				IEnumerator enumerator7 = default(IEnumerator);
				IEnumerator enumerator8 = default(IEnumerator);
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					try
					{
						enumerator5 = this.m_A.ActivePresentation.Slides.GetEnumerator();
						while (enumerator5.MoveNext())
						{
							Slide slide = (Slide)enumerator5.Current;
							foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
							{
								A(slide, shape);
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0333;
							}
							continue;
							end_IL_0333:
							break;
						}
					}
					finally
					{
						if (enumerator5 is IDisposable)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								(enumerator5 as IDisposable).Dispose();
								break;
							}
						}
					}
					if (chkLayouts.IsChecked != true)
					{
						break;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						try
						{
							enumerator7 = this.m_A.ActivePresentation.Designs.GetEnumerator();
							while (enumerator7.MoveNext())
							{
								Design design = (Design)enumerator7.Current;
								try
								{
									enumerator8 = design.SlideMaster.CustomLayouts.GetEnumerator();
									while (enumerator8.MoveNext())
									{
										CustomLayout customLayout = (CustomLayout)enumerator8.Current;
										foreach (Microsoft.Office.Interop.PowerPoint.Shape shape2 in customLayout.Shapes)
										{
											A(customLayout, shape2);
										}
									}
								}
								finally
								{
									if (enumerator8 is IDisposable)
									{
										while (true)
										{
											switch (5)
											{
											case 0:
												continue;
											}
											(enumerator8 as IDisposable).Dispose();
											break;
										}
									}
								}
							}
						}
						finally
						{
							if (enumerator7 is IDisposable)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									(enumerator7 as IDisposable).Dispose();
									break;
								}
							}
						}
						break;
					}
					break;
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.ErrorMessage(ex4.Message);
			ProjectData.ClearProjectError();
		}
		List<int> source;
		if (this.m_A.Any())
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
			source = this.m_A.Select([SpecialName] (Tuple<int, IndexedObject> A) => A.Item1).ToList();
			if (true)
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
				IEnumerable<A<int, int>> source2 = from A in source
					group A by A into A
					select new A<int, int>(A.Key, A.Count());
				Func<A<int, int>, int> keySelector;
				if (_Closure_0024__.A == null)
				{
					keySelector = (_Closure_0024__.A = [SpecialName] (A<int, int> A) => A.Count);
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
					keySelector = _Closure_0024__.A;
				}
				IOrderedEnumerable<A<int, int>> source3 = source2.OrderBy(keySelector);
				Func<A<int, int>, int> selector;
				if (_Closure_0024__.B == null)
				{
					selector = (_Closure_0024__.B = [SpecialName] (A<int, int> A) => A.Color);
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
					selector = _Closure_0024__.B;
				}
				source = source3.Select(selector).Reverse().ToList();
			}
			else
			{
				source = source.Distinct().ToList();
			}
			using List<int>.Enumerator enumerator10 = source.GetEnumerator();
			while (enumerator10.MoveNext())
			{
				int current2 = enumerator10.Current;
				System.Windows.Media.Color clr = A(ColorTranslator.FromOle(current2));
				this.m_A.Add(new ColorPair(clr, radRgb.IsChecked.Value));
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_0647;
				}
				continue;
				end_IL_0647:
				break;
			}
		}
		icColors.ItemsSource = this.m_A;
		selection = null;
		source = null;
	}

	private System.Windows.Media.Color A(System.Drawing.Color A)
	{
		return System.Windows.Media.Color.FromRgb(A.R, A.G, A.B);
	}

	private void ScopeChanged(object sender, RoutedEventArgs e)
	{
		System.Windows.Controls.CheckBox checkBox2;
		bool? obj;
		bool? isChecked2;
		System.Windows.Controls.CheckBox checkBox;
		if (sender is System.Windows.Controls.RadioButton)
		{
			bool? isChecked = ((System.Windows.Controls.RadioButton)sender).IsChecked;
			bool? flag;
			if (!isChecked.HasValue)
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
				flag = isChecked;
			}
			else
			{
				flag = isChecked != true;
			}
			isChecked = flag;
			if (isChecked == true)
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
				break;
			}
			checkBox = chkLayouts;
			checkBox2 = checkBox;
			bool? flag2 = (isChecked = radSlide.IsChecked);
			if (flag2.HasValue)
			{
				if (isChecked == true)
				{
					obj = true;
					goto IL_00f2;
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
			flag2 = (isChecked2 = radPresentation.IsChecked);
			if (!flag2.HasValue)
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
				obj = null;
			}
			else if (isChecked2 != true)
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
				obj = isChecked;
			}
			else
			{
				obj = true;
			}
			goto IL_00f2;
		}
		goto IL_0121;
		IL_00f2:
		isChecked2 = obj;
		checkBox2.IsEnabled = isChecked2.Value;
		if (!checkBox.IsEnabled)
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
			checkBox.IsChecked = false;
		}
		checkBox = null;
		goto IL_0121;
		IL_0121:
		C();
	}

	private void ColorDequeued(object sender, RoutedEventArgs e)
	{
		((ColorPair)((System.Windows.Controls.Button)sender).DataContext).Reset();
	}

	private void ChooseColor(object sender, RoutedEventArgs e)
	{
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Expected O, but got Unknown
		ColorPair colorPair = (ColorPair)((System.Windows.Controls.Button)sender).DataContext;
		wpfPalette val = new wpfPalette(false, (ColorRole)0);
		((Window)(object)val).WindowStartupLocation = WindowStartupLocation.CenterScreen;
		((Window)(object)val).Title = AH.A(12805);
		((Window)(object)val).ShowDialog();
		if (((Window)(object)val).DialogResult.HasValue)
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
			if (((Window)(object)val).DialogResult.Value)
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
				colorPair.NewColor = val.SelectedColor;
				btnRecolor.IsEnabled = true;
			}
		}
		val = null;
		colorPair = null;
	}

	private void ColorFormatChanged(object sender, RoutedEventArgs e)
	{
		bool value = radRgb.IsChecked.Value;
		IEnumerator<ColorPair> enumerator = default(IEnumerator<ColorPair>);
		try
		{
			enumerator = this.m_A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				enumerator.Current.GenerateLabels(value);
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
				return;
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

	private void btnRecolor_Click(object sender, RoutedEventArgs e)
	{
		bool value = chkFont.IsChecked.Value;
		bool value2 = chkFill.IsChecked.Value;
		bool value3 = chkBorder.IsChecked.Value;
		bool? isChecked = chkFont.IsChecked;
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
		bool? flag2 = flag;
		isChecked = flag;
		bool? obj;
		if (isChecked.HasValue)
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
			if (flag2 != true)
			{
				obj = false;
				goto IL_0145;
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
		isChecked = chkFill.IsChecked;
		bool? flag3;
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
			flag3 = isChecked;
		}
		else
		{
			flag3 = isChecked != true;
		}
		bool? flag4 = flag3;
		isChecked = flag3;
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
		else if (flag4 != true)
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
			obj = false;
		}
		else
		{
			obj = flag2;
		}
		goto IL_0145;
		IL_0145:
		bool? flag5 = obj;
		if (flag5.HasValue)
		{
			if (flag5 != true)
			{
				goto IL_01d0;
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
		flag4 = chkBorder.IsChecked;
		bool? flag6;
		if (!flag4.HasValue)
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
			flag6 = flag4;
		}
		else
		{
			flag6 = flag4 != true;
		}
		flag4 = flag6;
		if (flag4 == true && flag5.HasValue)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					Forms.WarningMessage(AH.A(12832));
					return;
				}
			}
		}
		goto IL_01d0;
		IL_0259:
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		try
		{
			List<ColorPair> list = this.m_A.Where([SpecialName] (ColorPair A) =>
			{
				if (A.NewColor.HasValue)
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
							return A.NewColor.Value != A.OldColor;
						}
					}
				}
				return false;
			}).ToList();
			using List<ColorPair>.Enumerator enumerator = list.GetEnumerator();
			while (enumerator.MoveNext())
			{
				ColorPair current = enumerator.Current;
				dictionary.Add(A(current.OldColor), A(current.NewColor.Value));
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_02eb;
				}
				continue;
				end_IL_02eb:
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		finally
		{
			List<ColorPair> list = null;
		}
		if (!dictionary.Any())
		{
			return;
		}
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
		IEnumerator enumerator4 = default(IEnumerator);
		IEnumerator enumerator5 = default(IEnumerator);
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			flag2 = (flag5 = chkLayouts.IsChecked);
			bool? obj2;
			if (flag2.HasValue)
			{
				if (flag5 != true)
				{
					obj2 = false;
					goto IL_0440;
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
			bool? flag7 = (flag2 = radSlide.IsChecked);
			bool? obj3;
			if (flag7.HasValue)
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
				if (flag2 == true)
				{
					obj3 = true;
					goto IL_03f3;
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
			flag7 = (isChecked = radPresentation.IsChecked);
			if (!flag7.HasValue)
			{
				obj3 = null;
			}
			else if (isChecked != true)
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
				obj3 = flag2;
			}
			else
			{
				obj3 = true;
			}
			goto IL_03f3;
			IL_0440:
			flag4 = obj2;
			bool value4 = flag4.Value;
			this.m_A.StartNewUndoEntry();
			if (value4)
			{
				if (radPresentation.IsChecked == true)
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
						enumerator2 = this.m_A.ActivePresentation.Designs.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Design design = (Design)enumerator2.Current;
							{
								enumerator3 = design.SlideMaster.CustomLayouts.GetEnumerator();
								try
								{
									while (enumerator3.MoveNext())
									{
										CustomLayout customLayout = (CustomLayout)enumerator3.Current;
										try
										{
											enumerator4 = customLayout.Shapes.GetEnumerator();
											while (enumerator4.MoveNext())
											{
												Microsoft.Office.Interop.PowerPoint.Shape a = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator4.Current;
												A(a, dictionary, value, value2, value3);
											}
											while (true)
											{
												switch (5)
												{
												case 0:
													break;
												default:
													goto end_IL_0516;
												}
												continue;
												end_IL_0516:
												break;
											}
										}
										finally
										{
											if (enumerator4 is IDisposable)
											{
												while (true)
												{
													switch (6)
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
									while (true)
									{
										switch (5)
										{
										case 0:
											break;
										default:
											goto end_IL_0550;
										}
										continue;
										end_IL_0550:
										break;
									}
								}
								finally
								{
									IDisposable disposable = enumerator3 as IDisposable;
									if (disposable != null)
									{
										disposable.Dispose();
									}
								}
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0580;
							}
							continue;
							end_IL_0580:
							break;
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (6)
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
				else if (radSlide.IsChecked == true)
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
						enumerator5 = this.m_A.ActiveWindow.Selection.SlideRange.GetEnumerator();
						while (enumerator5.MoveNext())
						{
							Slide slide = (Slide)enumerator5.Current;
							foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.CustomLayout.Shapes)
							{
								A(shape, dictionary, value, value2, value3);
							}
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_066c;
							}
							continue;
							end_IL_066c:
							break;
						}
					}
					finally
					{
						if (enumerator5 is IDisposable)
						{
							while (true)
							{
								switch (7)
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
			}
			using (List<Microsoft.Office.Interop.PowerPoint.Shape>.Enumerator enumerator7 = this.m_A.GetEnumerator())
			{
				while (enumerator7.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shape current2 = enumerator7.Current;
					if (current2.Type == MsoShapeType.msoPlaceholder)
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
						if (value4 && current2.HasChart == MsoTriState.msoFalse)
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
							if (current2.HasTable == MsoTriState.msoFalse)
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
								if (current2.HasSmartArt == MsoTriState.msoFalse)
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
									if (!value || current2.HasTextFrame != MsoTriState.msoTrue || current2.TextFrame2.HasText != MsoTriState.msoTrue)
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
									this.m_A = new List<object>();
									using (Dictionary<int, int>.Enumerator enumerator8 = dictionary.GetEnumerator())
									{
										while (enumerator8.MoveNext())
										{
											KeyValuePair<int, int> current3 = enumerator8.Current;
											TextRange2 textRange = current2.TextFrame2.TextRange;
											int count = textRange.get_Runs(-1, -1).Count;
											for (int num = 1; num <= count; num = checked(num + 1))
											{
												TextRange2 textRange2 = textRange.get_Runs(num, -1);
												A(textRange2.Font, current3.Key, current3.Value);
												B(textRange2.Font, current3.Key, current3.Value);
												C(textRange2.Font, current3.Key, current3.Value);
												textRange2 = null;
											}
											textRange = null;
										}
										while (true)
										{
											switch (7)
											{
											case 0:
												break;
											default:
												goto end_IL_0823;
											}
											continue;
											end_IL_0823:
											break;
										}
									}
									continue;
								}
							}
						}
					}
					A(current2, dictionary, value, value2, value3);
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_0858;
					}
					continue;
					end_IL_0858:
					break;
				}
			}
			C();
			clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)1, AH.A(12720));
			return;
			IL_03f3:
			flag4 = obj3;
			isChecked = obj3;
			if (!isChecked.HasValue)
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
				obj2 = null;
			}
			else if (flag4 != true)
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
				obj2 = false;
			}
			else
			{
				obj2 = flag5;
			}
			goto IL_0440;
		}
		IL_01d0:
		flag5 = radSlide.IsChecked;
		if (flag5.HasValue)
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
			if (flag5 != true)
			{
				goto IL_0259;
			}
		}
		if (chkLayouts.IsChecked == true)
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
			if (flag5.HasValue)
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
				if (Forms.OkCancelMessage(AH.A(12889)) == DialogResult.Cancel)
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
					break;
				}
			}
		}
		goto IL_0259;
	}

	private int A(System.Windows.Media.Color A)
	{
		return ColorTranslator.ToOle(System.Drawing.Color.FromArgb(A.R, A.G, A.B));
	}

	private void btnRefresh_Click(object sender, RoutedEventArgs e)
	{
		C();
	}

	private void A(object A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		if (Base.A(B))
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
			Microsoft.Office.Interop.PowerPoint.Shape shape = B;
			if (shape.HasTable == MsoTriState.msoTrue)
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
				this.m_A.Add(B);
				this.B(RuntimeHelpers.GetObjectValue(A), B);
			}
			else if (shape.HasChart == MsoTriState.msoTrue)
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
				this.m_A.Add(B);
				D(RuntimeHelpers.GetObjectValue(A), B);
			}
			else if (shape.HasSmartArt == MsoTriState.msoTrue)
			{
				this.m_A.Add(B);
				C(RuntimeHelpers.GetObjectValue(A), B);
			}
			else if (shape.Type == MsoShapeType.msoGroup)
			{
				int count = shape.GroupItems.Count;
				for (int i = 1; i <= count; i = checked(i + 1))
				{
					this.A(RuntimeHelpers.GetObjectValue(A), shape.GroupItems[i]);
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
				this.m_A.Add(B);
				if (chkFill.IsChecked == true)
				{
					Index.Fill(RuntimeHelpers.GetObjectValue(A), B, B, ref this.m_A);
				}
			}
			else
			{
				this.m_A.Add(B);
				if (chkFont.IsChecked == true)
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
					Index.Font(RuntimeHelpers.GetObjectValue(A), B, B, ref this.m_A);
				}
				if (chkFill.IsChecked == true)
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
					Index.Fill(RuntimeHelpers.GetObjectValue(A), B, B, ref this.m_A);
				}
				if (chkBorder.IsChecked == true)
				{
					Index.Border(RuntimeHelpers.GetObjectValue(A), B, B, ref this.m_A);
				}
			}
			shape = null;
			return;
		}
	}

	private void B(object A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		List<Tuple<int, IndexedObject>> FontColors = new List<Tuple<int, IndexedObject>>();
		List<Tuple<int, IndexedObject>> FillColors = new List<Tuple<int, IndexedObject>>();
		List<Tuple<int, IndexedObject>> BorderColors = new List<Tuple<int, IndexedObject>>();
		Index.Table(RuntimeHelpers.GetObjectValue(A), B, B.Table, chkFont.IsChecked.Value, chkFill.IsChecked.Value, chkBorder.IsChecked.Value, ref FontColors, ref FillColors, ref BorderColors);
		this.A(FontColors, FillColors, BorderColors);
		FontColors = null;
		FillColors = null;
		BorderColors = null;
	}

	private void C(object A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		List<Tuple<int, IndexedObject>> FontColors = new List<Tuple<int, IndexedObject>>();
		List<Tuple<int, IndexedObject>> FillColors = new List<Tuple<int, IndexedObject>>();
		List<Tuple<int, IndexedObject>> BorderColors = new List<Tuple<int, IndexedObject>>();
		Index.SmartArt(RuntimeHelpers.GetObjectValue(A), B, B.SmartArt, chkFont.IsChecked.Value, chkFill.IsChecked.Value, chkBorder.IsChecked.Value, ref FontColors, ref FillColors, ref BorderColors);
		this.A(FontColors, FillColors, BorderColors);
		FontColors = null;
		FillColors = null;
		BorderColors = null;
	}

	private void D(object A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		List<Tuple<int, IndexedObject>> FontColors = new List<Tuple<int, IndexedObject>>();
		List<Tuple<int, IndexedObject>> FillColors = new List<Tuple<int, IndexedObject>>();
		List<Tuple<int, IndexedObject>> BorderColors = new List<Tuple<int, IndexedObject>>();
		Index.Chart(RuntimeHelpers.GetObjectValue(A), B, B.Chart, chkFont.IsChecked.Value, chkFill.IsChecked.Value, chkBorder.IsChecked.Value, ref FontColors, ref FillColors, ref BorderColors);
		this.A(FontColors, FillColors, BorderColors);
		FontColors = null;
		FillColors = null;
		BorderColors = null;
	}

	private void A(List<Tuple<int, IndexedObject>> A, List<Tuple<int, IndexedObject>> B, List<Tuple<int, IndexedObject>> C)
	{
		List<Tuple<int, IndexedObject>> a = this.m_A;
		a.AddRange(A);
		a.AddRange(B);
		a.AddRange(C);
		_ = null;
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A, Dictionary<int, int> B, bool C, bool D, bool E)
	{
		this.m_A = new List<object>();
		try
		{
			if (A.HasChart == MsoTriState.msoTrue)
			{
				using (Dictionary<int, int>.Enumerator enumerator = B.GetEnumerator())
				{
					while (enumerator.MoveNext())
					{
						KeyValuePair<int, int> current = enumerator.Current;
						this.A(A.Chart, current.Key, current.Value, C, D, E);
					}
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
							return;
						}
					}
				}
			}
			if (A.HasTable == MsoTriState.msoTrue)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
					{
						using Dictionary<int, int>.Enumerator enumerator2 = B.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							KeyValuePair<int, int> current2 = enumerator2.Current;
							this.A(A.Table, current2.Key, current2.Value, C, D, E);
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
					}
				}
			}
			if (A.HasSmartArt == MsoTriState.msoTrue)
			{
				using (Dictionary<int, int>.Enumerator enumerator3 = B.GetEnumerator())
				{
					while (enumerator3.MoveNext())
					{
						KeyValuePair<int, int> current3 = enumerator3.Current;
						this.A(A, current3.Key, current3.Value, C, D, E);
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
			}
			IEnumerator enumerator4 = default(IEnumerator);
			if (A.Type == MsoShapeType.msoGroup)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						if (D)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
								{
									Dictionary<Microsoft.Office.Interop.PowerPoint.Shape, Tuple<int, int>> dictionary = new Dictionary<Microsoft.Office.Interop.PowerPoint.Shape, Tuple<int, int>>();
									try
									{
										enumerator4 = A.GroupItems.GetEnumerator();
										while (enumerator4.MoveNext())
										{
											Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator4.Current;
											Microsoft.Office.Interop.PowerPoint.FillFormat fill = shape.Fill;
											if (fill.Visible == MsoTriState.msoTrue)
											{
												dictionary.Add(shape, new Tuple<int, int>(fill.ForeColor.RGB, fill.BackColor.RGB));
											}
											fill = null;
										}
									}
									finally
									{
										if (enumerator4 is IDisposable)
										{
											while (true)
											{
												switch (7)
												{
												case 0:
													break;
												default:
													(enumerator4 as IDisposable).Dispose();
													goto end_IL_0201;
												}
												continue;
												end_IL_0201:
												break;
											}
										}
									}
									using (Dictionary<int, int>.Enumerator enumerator5 = B.GetEnumerator())
									{
										while (enumerator5.MoveNext())
										{
											KeyValuePair<int, int> current4 = enumerator5.Current;
											this.B(A, current4.Key, current4.Value);
										}
										while (true)
										{
											switch (4)
											{
											case 0:
												break;
											default:
												goto end_IL_0252;
											}
											continue;
											end_IL_0252:
											break;
										}
									}
									using (Dictionary<Microsoft.Office.Interop.PowerPoint.Shape, Tuple<int, int>>.Enumerator enumerator6 = dictionary.GetEnumerator())
									{
										while (enumerator6.MoveNext())
										{
											KeyValuePair<Microsoft.Office.Interop.PowerPoint.Shape, Tuple<int, int>> current5 = enumerator6.Current;
											Microsoft.Office.Interop.PowerPoint.FillFormat fill2 = current5.Key.Fill;
											fill2.ForeColor.RGB = current5.Value.Item1;
											fill2.BackColor.RGB = current5.Value.Item2;
											_ = null;
										}
										while (true)
										{
											switch (6)
											{
											case 0:
												break;
											default:
												goto end_IL_02d5;
											}
											continue;
											end_IL_02d5:
											break;
										}
									}
									dictionary = null;
									return;
								}
								}
							}
						}
						return;
					}
				}
			}
			using Dictionary<int, int>.Enumerator enumerator7 = B.GetEnumerator();
			while (enumerator7.MoveNext())
			{
				KeyValuePair<int, int> current6 = enumerator7.Current;
				if (C)
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
					this.A(A, current6.Key, current6.Value);
				}
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
					this.B(A, current6.Key, current6.Value);
				}
				if (!E)
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
				this.C(A, current6.Key, current6.Value);
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A, int B, int C)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape = A;
			if (shape.HasTextFrame == MsoTriState.msoTrue)
			{
				if (shape.TextFrame2.HasText == MsoTriState.msoTrue)
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
					TextRange2 textRange = shape.TextFrame2.TextRange;
					int count = textRange.get_Runs(-1, -1).Count;
					for (int i = 1; i <= count; i = checked(i + 1))
					{
						this.A(textRange.get_Runs(i, -1).Font, B, C);
						this.B(textRange.get_Runs(i, -1).Font, B, C);
						this.C(textRange.get_Runs(i, -1).Font, B, C);
						Microsoft.Office.Core.FillFormat fill = textRange.get_Runs(i, -1).Font.Fill;
						this.A(fill.ForeColor, B, C);
						this.A(fill.BackColor, B, C);
						fill = null;
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
					textRange = null;
				}
				else
				{
					this.A(shape.TextFrame2.TextRange.Font, B, C);
					this.B(shape.TextFrame2.TextRange.Font, B, C);
					this.C(shape.TextFrame2.TextRange.Font, B, C);
					Microsoft.Office.Core.FillFormat fill2 = shape.TextFrame2.TextRange.Font.Fill;
					this.A(fill2.ForeColor, B, C);
					this.A(fill2.BackColor, B, C);
					fill2 = null;
				}
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = shape.TextFrame2.TextRange.get_Paragraphs(-1, -1).GetEnumerator();
					while (enumerator.MoveNext())
					{
						BulletFormat2 bullet = ((TextRange2)enumerator.Current).ParagraphFormat.Bullet;
						if (bullet.Type == MsoBulletType.msoBulletNumbered || bullet.Type == MsoBulletType.msoBulletUnnumbered)
						{
							Microsoft.Office.Core.FillFormat fill3 = bullet.Font.Fill;
							this.A(fill3.ForeColor, B, C);
							this.A(fill3.BackColor, B, C);
							fill3 = null;
						}
						bullet = null;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_0226;
						}
						continue;
						end_IL_0226:
						break;
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
			shape = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void A(Font2 A, int B, int C)
	{
		try
		{
			Font2 font = A;
			if (font.UnderlineStyle != MsoTextUnderlineType.msoNoUnderline)
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
				if (font.UnderlineColor.RGB == B && !this.m_A.Contains(font.UnderlineColor))
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
					float size = font.Size;
					string name = font.Name;
					int rGB = font.Fill.ForeColor.RGB;
					font.UnderlineColor.RGB = C;
					this.m_A.Add(font.UnderlineColor);
					font.Size = size;
					font.Name = name;
					font.Fill.ForeColor.RGB = rGB;
				}
			}
			font = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void B(Font2 A, int B, int C)
	{
		try
		{
			Microsoft.Office.Core.LineFormat line = A.Line;
			if (line.Visible == MsoTriState.msoTrue)
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
				this.A(line.ForeColor, B, C);
			}
			line = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void C(Font2 A, int B, int C)
	{
		try
		{
			Font2 font = A;
			if (B == 0)
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
				if (font.Highlight.Type != MsoColorType.msoColorTypeRGB)
				{
					goto IL_00ce;
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
			if (font.Highlight.RGB == B)
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
				if (!this.m_A.Contains(font.Highlight))
				{
					float size = font.Size;
					string name = font.Name;
					int rGB = font.Fill.ForeColor.RGB;
					font.Highlight.RGB = C;
					this.m_A.Add(font.Highlight);
					font.Size = size;
					font.Name = name;
					font.Fill.ForeColor.RGB = rGB;
				}
			}
			goto IL_00ce;
			IL_00ce:
			font = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void B(Microsoft.Office.Interop.PowerPoint.Shape A, int B, int C)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.FillFormat fill = A.Fill;
			if (fill.Visible == MsoTriState.msoTrue)
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
				this.A(fill.ForeColor, B, C);
				this.A(fill.BackColor, B, C);
			}
			fill = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void C(Microsoft.Office.Interop.PowerPoint.Shape A, int B, int C)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.LineFormat line = A.Line;
			if (line.Visible == MsoTriState.msoTrue)
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
				this.A(line.ForeColor, B, C);
			}
			line = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void A(Microsoft.Office.Core.ColorFormat A, int B, int C)
	{
		if (A.RGB != B)
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
			if (this.m_A.Contains(A))
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
				A.RGB = C;
				this.m_A.Add(A);
				return;
			}
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.ColorFormat A, int B, int C)
	{
		if (A.RGB != B)
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
			if (!this.m_A.Contains(A))
			{
				A.RGB = C;
				this.m_A.Add(A);
			}
			return;
		}
	}

	private void A(Table A, int B, int C, bool D, bool E, bool F)
	{
		int count = A.Rows.Count;
		int count2 = A.Columns.Count;
		int num = count;
		checked
		{
			for (int i = 1; i <= num; i++)
			{
				int num2 = count2;
				for (int j = 1; j <= num2; j++)
				{
					Cell cell = A.Cell(i, j);
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						this.A(cell.Shape, B, C);
					}
					if (E)
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
						this.B(cell.Shape, B, C);
					}
					if (F)
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
						this.C(cell.Shape, B, C);
					}
					cell = null;
				}
			}
		}
	}

	private void A(Chart A, int B, int C, bool D, bool E, bool F)
	{
		bool flag = clsCharts.UsesLegendsForSeriesClrs(A);
		bool d = flag && clsCharts.UsesLegendLinesForSeriesClrs(A);
		bool flag2 = clsCharts.UsesFormatFillForSeriesClrs(A);
		if (E)
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
				if (A.HasLegend)
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
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = ((IEnumerable)A.Legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						while (enumerator.MoveNext())
						{
							object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
							try
							{
								if (NewLateBinding.LateGet(objectValue, null, AH.A(13177), new object[0], null, null, null) is IMsoLegendKey a)
								{
									while (true)
									{
										switch (6)
										{
										case 0:
											continue;
										}
										this.A(a, B, C, d);
										break;
									}
								}
							}
							catch (Exception projectError)
							{
								ProjectData.SetProjectError(projectError);
								ProjectData.ClearProjectError();
							}
							IMsoLegendKey msoLegendKey = null;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_00fc;
							}
							continue;
							end_IL_00fc:
							break;
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
			}
		}
		if (!clsCharts.SeriesClrsAreUnusable(A))
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
			IEnumerator enumerator2 = default(IEnumerator);
			try
			{
				enumerator2 = ((IEnumerable)A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
				IEnumerator enumerator3 = default(IEnumerator);
				int rGB = default(int);
				int rGB2 = default(int);
				int rGB3 = default(int);
				int rGB4 = default(int);
				int markerForegroundColor = default(int);
				int markerBackgroundColor = default(int);
				IEnumerator enumerator4 = default(IEnumerator);
				while (enumerator2.MoveNext())
				{
					IMsoSeries msoSeries = (IMsoSeries)enumerator2.Current;
					if (F)
					{
						if (Charts.ImplsAndHasErrorBars(msoSeries))
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
							this.A(msoSeries.ErrorBars.Format.Line, B, C);
						}
						if (msoSeries.HasLeaderLines)
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
							this.A(msoSeries.LeaderLines.Format.Line, B, C);
						}
						if (Charts.ImplsTrendLines(msoSeries))
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
								enumerator3 = ((IEnumerable)msoSeries.Trendlines(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
								while (enumerator3.MoveNext())
								{
									IMsoTrendline msoTrendline = (IMsoTrendline)enumerator3.Current;
									this.A(msoTrendline.Format.Line, B, C);
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										goto end_IL_0245;
									}
									continue;
									end_IL_0245:
									break;
								}
							}
							finally
							{
								if (enumerator3 is IDisposable)
								{
									while (true)
									{
										switch (2)
										{
										case 0:
											continue;
										}
										(enumerator3 as IDisposable).Dispose();
										break;
									}
								}
							}
						}
					}
					Dictionary<ChartPoint, int> dictionary = new Dictionary<ChartPoint, int>();
					Dictionary<ChartPoint, int> dictionary2 = new Dictionary<ChartPoint, int>();
					Dictionary<ChartPoint, int> dictionary3 = new Dictionary<ChartPoint, int>();
					Dictionary<ChartPoint, int> dictionary4 = new Dictionary<ChartPoint, int>();
					Dictionary<ChartPoint, int> dictionary5 = new Dictionary<ChartPoint, int>();
					Dictionary<ChartPoint, int> dictionary6 = new Dictionary<ChartPoint, int>();
					if (flag2)
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
						Microsoft.Office.Core.FillFormat fill = msoSeries.Format.Fill;
						rGB = fill.ForeColor.RGB;
						rGB2 = fill.BackColor.RGB;
						_ = null;
						Microsoft.Office.Core.LineFormat line = msoSeries.Format.Line;
						rGB3 = line.ForeColor.RGB;
						rGB4 = line.BackColor.RGB;
						_ = null;
					}
					bool flag3;
					try
					{
						if (msoSeries.MarkerStyle != XlMarkerStyle.xlMarkerStyleNone)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								markerForegroundColor = msoSeries.MarkerForegroundColor;
								markerBackgroundColor = msoSeries.MarkerBackgroundColor;
								flag3 = true;
								break;
							}
						}
						else
						{
							flag3 = false;
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						flag3 = false;
						ProjectData.ClearProjectError();
					}
					if (Charts.ImplsPoints(msoSeries))
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
							enumerator4 = ((IEnumerable)msoSeries.Points(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
							while (enumerator4.MoveNext())
							{
								ChartPoint chartPoint = (ChartPoint)enumerator4.Current;
								if (E)
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
										Microsoft.Office.Core.ColorFormat foreColor = chartPoint.Format.Fill.ForeColor;
										if (foreColor.RGB == rGB)
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
											if (flag2)
											{
												goto IL_0400;
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
										dictionary3.Add(chartPoint, foreColor.RGB);
										goto IL_0400;
										IL_0400:
										foreColor = null;
									}
									catch (Exception ex3)
									{
										ProjectData.SetProjectError(ex3);
										Exception ex4 = ex3;
										ProjectData.ClearProjectError();
									}
									try
									{
										Microsoft.Office.Core.ColorFormat backColor = chartPoint.Format.Fill.BackColor;
										if (backColor.RGB == rGB2)
										{
											if (flag2)
											{
												goto IL_0456;
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
										dictionary4.Add(chartPoint, backColor.RGB);
										goto IL_0456;
										IL_0456:
										backColor = null;
									}
									catch (Exception ex5)
									{
										ProjectData.SetProjectError(ex5);
										Exception ex6 = ex5;
										ProjectData.ClearProjectError();
									}
								}
								if (F)
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
									try
									{
										Microsoft.Office.Core.ColorFormat foreColor2 = chartPoint.Format.Line.ForeColor;
										if (foreColor2.RGB != rGB3)
										{
											dictionary5.Add(chartPoint, foreColor2.RGB);
										}
										foreColor2 = null;
									}
									catch (Exception ex7)
									{
										ProjectData.SetProjectError(ex7);
										Exception ex8 = ex7;
										ProjectData.ClearProjectError();
									}
									try
									{
										Microsoft.Office.Core.ColorFormat backColor2 = chartPoint.Format.Line.BackColor;
										if (backColor2.RGB != rGB4)
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
											dictionary6.Add(chartPoint, backColor2.RGB);
										}
										backColor2 = null;
									}
									catch (Exception ex9)
									{
										ProjectData.SetProjectError(ex9);
										Exception ex10 = ex9;
										ProjectData.ClearProjectError();
									}
								}
								bool flag4 = false;
								try
								{
									int num;
									if (chartPoint.MarkerStyle != XlMarkerStyle.xlMarkerStyleNone)
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
										num = ((chartPoint.MarkerStyle != XlMarkerStyle.xlMarkerStylePicture) ? 1 : 0);
									}
									else
									{
										num = 0;
									}
									flag4 = (byte)num != 0;
								}
								catch (Exception projectError2)
								{
									ProjectData.SetProjectError(projectError2);
									ProjectData.ClearProjectError();
								}
								if (flag4)
								{
									if (E)
									{
										try
										{
											if (chartPoint.MarkerBackgroundColor != markerBackgroundColor)
											{
												dictionary2.Add(chartPoint, chartPoint.MarkerBackgroundColor);
											}
										}
										catch (Exception ex11)
										{
											ProjectData.SetProjectError(ex11);
											Exception ex12 = ex11;
											ProjectData.ClearProjectError();
										}
									}
									if (F)
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
											if (chartPoint.MarkerForegroundColor != markerForegroundColor)
											{
												while (true)
												{
													switch (1)
													{
													case 0:
														continue;
													}
													dictionary.Add(chartPoint, chartPoint.MarkerForegroundColor);
													break;
												}
											}
										}
										catch (Exception ex13)
										{
											ProjectData.SetProjectError(ex13);
											Exception ex14 = ex13;
											ProjectData.ClearProjectError();
										}
									}
								}
								if (!D)
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
								try
								{
									this.A(chartPoint.DataLabel.Font, B, C);
								}
								catch (Exception ex15)
								{
									ProjectData.SetProjectError(ex15);
									Exception ex16 = ex15;
									ProjectData.ClearProjectError();
								}
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_061f;
								}
								continue;
								end_IL_061f:
								break;
							}
						}
						finally
						{
							if (enumerator4 is IDisposable)
							{
								while (true)
								{
									switch (1)
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
					if (E)
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
							this.A(msoSeries.Format.Fill, B, C);
							using (Dictionary<ChartPoint, int>.Enumerator enumerator5 = dictionary3.GetEnumerator())
							{
								while (enumerator5.MoveNext())
								{
									KeyValuePair<ChartPoint, int> current = enumerator5.Current;
									Microsoft.Office.Core.FillFormat fill2 = current.Key.Format.Fill;
									if (fill2.Visible == MsoTriState.msoTrue)
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
										Microsoft.Office.Core.ColorFormat foreColor3 = fill2.ForeColor;
										if (current.Value != B)
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
											foreColor3.RGB = current.Value;
										}
										else if (foreColor3.RGB == B)
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
											foreColor3.RGB = C;
										}
										foreColor3 = null;
									}
									fill2 = null;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_071e;
									}
									continue;
									end_IL_071e:
									break;
								}
							}
							using Dictionary<ChartPoint, int>.Enumerator enumerator6 = dictionary4.GetEnumerator();
							while (enumerator6.MoveNext())
							{
								KeyValuePair<ChartPoint, int> current2 = enumerator6.Current;
								Microsoft.Office.Core.FillFormat fill3 = current2.Key.Format.Fill;
								if (fill3.Visible == MsoTriState.msoTrue)
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
									Microsoft.Office.Core.ColorFormat backColor3 = fill3.BackColor;
									if (current2.Value != B)
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
										backColor3.RGB = current2.Value;
									}
									else if (backColor3.RGB == B)
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
										backColor3.RGB = C;
									}
									backColor3 = null;
								}
								fill3 = null;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_07e3;
								}
								continue;
								end_IL_07e3:
								break;
							}
						}
						catch (Exception ex17)
						{
							ProjectData.SetProjectError(ex17);
							Exception ex18 = ex17;
							ProjectData.ClearProjectError();
						}
					}
					dictionary3 = null;
					dictionary4 = null;
					if (F)
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
						try
						{
							if (msoSeries.Format.Line.Visible == MsoTriState.msoTrue)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									if (!flag)
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
										this.A(msoSeries.Format.Line, B, C);
									}
									using (Dictionary<ChartPoint, int>.Enumerator enumerator7 = dictionary5.GetEnumerator())
									{
										while (enumerator7.MoveNext())
										{
											KeyValuePair<ChartPoint, int> current3 = enumerator7.Current;
											Microsoft.Office.Core.LineFormat line2 = current3.Key.Format.Line;
											if (line2.Visible == MsoTriState.msoTrue)
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
												Microsoft.Office.Core.ColorFormat foreColor4 = line2.ForeColor;
												if (current3.Value != B)
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
													foreColor4.RGB = current3.Value;
												}
												else if (foreColor4.RGB == B)
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
													foreColor4.RGB = C;
												}
												foreColor4 = null;
											}
											line2 = null;
										}
										while (true)
										{
											switch (7)
											{
											case 0:
												break;
											default:
												goto end_IL_0918;
											}
											continue;
											end_IL_0918:
											break;
										}
									}
									using (Dictionary<ChartPoint, int>.Enumerator enumerator8 = dictionary6.GetEnumerator())
									{
										while (enumerator8.MoveNext())
										{
											KeyValuePair<ChartPoint, int> current4 = enumerator8.Current;
											Microsoft.Office.Core.LineFormat line3 = current4.Key.Format.Line;
											if (line3.Visible == MsoTriState.msoTrue)
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
												Microsoft.Office.Core.ColorFormat backColor4 = line3.BackColor;
												if (current4.Value != B)
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
													backColor4.RGB = current4.Value;
												}
												else if (backColor4.RGB == B)
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
													backColor4.RGB = C;
												}
												backColor4 = null;
											}
											line3 = null;
										}
										while (true)
										{
											switch (1)
											{
											case 0:
												break;
											default:
												goto end_IL_09de;
											}
											continue;
											end_IL_09de:
											break;
										}
									}
									break;
								}
							}
						}
						catch (Exception ex19)
						{
							ProjectData.SetProjectError(ex19);
							Exception ex20 = ex19;
							ProjectData.ClearProjectError();
						}
					}
					dictionary5 = null;
					dictionary6 = null;
					if (flag3)
					{
						if (E)
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
							if (msoSeries.MarkerBackgroundColor == B)
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
								msoSeries.MarkerBackgroundColor = C;
							}
						}
						if (F)
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
							if (msoSeries.MarkerForegroundColor == B)
							{
								msoSeries.MarkerForegroundColor = C;
							}
						}
						foreach (KeyValuePair<ChartPoint, int> item in dictionary2)
						{
							if (item.Key.MarkerBackgroundColor == B)
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
								item.Key.MarkerBackgroundColor = C;
							}
							else
							{
								item.Key.MarkerBackgroundColor = item.Value;
							}
						}
						foreach (KeyValuePair<ChartPoint, int> item2 in dictionary)
						{
							if (item2.Key.MarkerForegroundColor == B)
							{
								item2.Key.MarkerForegroundColor = C;
							}
							else
							{
								item2.Key.MarkerForegroundColor = item2.Value;
							}
						}
					}
					dictionary = null;
					dictionary2 = null;
				}
			}
			finally
			{
				if (enumerator2 is IDisposable)
				{
					while (true)
					{
						switch (2)
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
		int count = ((ChartGroups)A.ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value))).Count;
		for (int i = 1; i <= count; i = checked(i + 1))
		{
			ChartGroup chartGroup = (ChartGroup)A.ChartGroups(i);
			if (chartGroup.HasUpDownBars)
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
				if (E)
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
					try
					{
						this.A(chartGroup.UpBars.Format.Fill, B, C);
						this.A(chartGroup.DownBars.Format.Fill, B, C);
					}
					catch (Exception ex21)
					{
						ProjectData.SetProjectError(ex21);
						Exception ex22 = ex21;
						ProjectData.ClearProjectError();
					}
				}
				if (F)
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
						this.A(chartGroup.UpBars.Format.Line, B, C);
						this.A(chartGroup.DownBars.Format.Line, B, C);
					}
					catch (Exception ex23)
					{
						ProjectData.SetProjectError(ex23);
						Exception ex24 = ex23;
						ProjectData.ClearProjectError();
					}
				}
			}
			if (chartGroup.HasHiLoLines)
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
				if (F)
				{
					try
					{
						this.A(chartGroup.HiLoLines.Format.Line, B, C);
					}
					catch (Exception ex25)
					{
						ProjectData.SetProjectError(ex25);
						Exception ex26 = ex25;
						ProjectData.ClearProjectError();
					}
				}
			}
			if (chartGroup.HasDropLines && F)
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
				try
				{
					this.A(chartGroup.DropLines.Format.Line, B, C);
				}
				catch (Exception ex27)
				{
					ProjectData.SetProjectError(ex27);
					Exception ex28 = ex27;
					ProjectData.ClearProjectError();
				}
			}
			if (Charts.HasRadarAxisLabels(chartGroup))
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
					this.A(chartGroup.RadarAxisLabels.Font, B, C);
				}
				catch (Exception projectError3)
				{
					ProjectData.SetProjectError(projectError3);
					ProjectData.ClearProjectError();
				}
			}
			chartGroup = null;
		}
		IEnumerator enumerator11 = default(IEnumerator);
		IEnumerator enumerator12 = default(IEnumerator);
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (A.HasTitle)
			{
				ChartTitle chartTitle = A.ChartTitle;
				if (D)
				{
					try
					{
						if (clsCharts.ImplsFont(A.ChartTitle))
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								this.A(chartTitle.Font, B, C);
								break;
							}
						}
						else if (B != 0)
						{
							this.A(chartTitle.Format.TextFrame2.TextRange, B, C);
						}
					}
					catch (Exception ex29)
					{
						ProjectData.SetProjectError(ex29);
						Exception ex30 = ex29;
						ProjectData.ClearProjectError();
					}
				}
				if (E)
				{
					try
					{
						this.A(chartTitle.Format.Fill, B, C);
					}
					catch (Exception ex31)
					{
						ProjectData.SetProjectError(ex31);
						Exception ex32 = ex31;
						ProjectData.ClearProjectError();
					}
				}
				if (F)
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
						this.A(chartTitle.Format.Line, B, C);
					}
					catch (Exception ex33)
					{
						ProjectData.SetProjectError(ex33);
						Exception ex34 = ex33;
						ProjectData.ClearProjectError();
					}
				}
				chartTitle = null;
			}
			if (A.HasLegend)
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
				Legend legend = A.Legend;
				if (D)
				{
					try
					{
						List<Tuple<Microsoft.Office.Core.LegendEntry, int>> list = new List<Tuple<Microsoft.Office.Core.LegendEntry, int>>();
						Microsoft.Office.Core.ColorFormat foreColor5 = legend.Format.TextFrame2.TextRange.get_Characters(-1, -1).Font.Fill.ForeColor;
						if (foreColor5.RGB == B)
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
							if (B != C)
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
									enumerator11 = ((IEnumerable)A.Legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
									while (enumerator11.MoveNext())
									{
										Microsoft.Office.Core.LegendEntry legendEntry = (Microsoft.Office.Core.LegendEntry)enumerator11.Current;
										if (!Charts.ImplsFont(legendEntry) || !Operators.ConditionalCompareObjectNotEqual(legendEntry.Font.Color, B, TextCompare: false))
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
										if (!Operators.ConditionalCompareObjectNotEqual(legendEntry.Font.Color, C, TextCompare: false))
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
										list.Add(new Tuple<Microsoft.Office.Core.LegendEntry, int>(legendEntry, Conversions.ToInteger(legendEntry.Font.Color)));
									}
								}
								finally
								{
									if (enumerator11 is IDisposable)
									{
										while (true)
										{
											switch (1)
											{
											case 0:
												continue;
											}
											(enumerator11 as IDisposable).Dispose();
											break;
										}
									}
								}
								foreColor5.RGB = C;
							}
						}
						foreColor5 = null;
						try
						{
							enumerator12 = ((IEnumerable)legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
							while (enumerator12.MoveNext())
							{
								Microsoft.Office.Core.LegendEntry legendEntry2 = (Microsoft.Office.Core.LegendEntry)enumerator12.Current;
								if (!Charts.ImplsFont(legendEntry2))
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
								this.A(legendEntry2.Font, B, C);
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_1027;
								}
								continue;
								end_IL_1027:
								break;
							}
						}
						finally
						{
							if (enumerator12 is IDisposable)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									(enumerator12 as IDisposable).Dispose();
									break;
								}
							}
						}
						using List<Tuple<Microsoft.Office.Core.LegendEntry, int>>.Enumerator enumerator13 = list.GetEnumerator();
						while (enumerator13.MoveNext())
						{
							Tuple<Microsoft.Office.Core.LegendEntry, int> current7 = enumerator13.Current;
							if (Charts.ImplsFont(current7.Item1))
							{
								current7.Item1.Font.Color = current7.Item2;
							}
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_10aa;
							}
							continue;
							end_IL_10aa:
							break;
						}
					}
					catch (Exception ex35)
					{
						ProjectData.SetProjectError(ex35);
						Exception ex36 = ex35;
						ProjectData.ClearProjectError();
					}
				}
				if (flag2)
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
					if (E)
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
							this.A(legend.Format.Fill, B, C);
						}
						catch (Exception ex37)
						{
							ProjectData.SetProjectError(ex37);
							Exception ex38 = ex37;
							ProjectData.ClearProjectError();
						}
					}
					if (F)
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
							this.A(legend.Format.Line, B, C);
						}
						catch (Exception ex39)
						{
							ProjectData.SetProjectError(ex39);
							Exception ex40 = ex39;
							ProjectData.ClearProjectError();
						}
					}
				}
				legend = null;
			}
			if (A.HasDataTable)
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
				DataTable dataTable = A.DataTable;
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
					try
					{
						this.A(dataTable.Font, B, C);
					}
					catch (Exception ex41)
					{
						ProjectData.SetProjectError(ex41);
						Exception ex42 = ex41;
						ProjectData.ClearProjectError();
					}
				}
				if (E)
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
					try
					{
						this.A(dataTable.Format.Fill, B, C);
					}
					catch (Exception ex43)
					{
						ProjectData.SetProjectError(ex43);
						Exception ex44 = ex43;
						ProjectData.ClearProjectError();
					}
				}
				if (F)
				{
					try
					{
						this.A(dataTable.Format.Line, B, C);
					}
					catch (Exception ex45)
					{
						ProjectData.SetProjectError(ex45);
						Exception ex46 = ex45;
						ProjectData.ClearProjectError();
					}
				}
				dataTable = null;
			}
			foreach (Axis item3 in modCharts.AxesList(A))
			{
				try
				{
					this.A(A, item3, D, E, F, B, C);
				}
				finally
				{
					Axis current8 = null;
				}
			}
			if (E)
			{
				try
				{
					this.A(A.ChartArea.Format.Fill, B, C);
				}
				catch (Exception ex47)
				{
					ProjectData.SetProjectError(ex47);
					Exception ex48 = ex47;
					ProjectData.ClearProjectError();
				}
				try
				{
					this.A(A.PlotArea.Format.Fill, B, C);
				}
				catch (Exception ex49)
				{
					ProjectData.SetProjectError(ex49);
					Exception ex50 = ex49;
					ProjectData.ClearProjectError();
				}
			}
			if (F)
			{
				try
				{
					this.A(A.ChartArea.Format.Line, B, C);
				}
				catch (Exception ex51)
				{
					ProjectData.SetProjectError(ex51);
					Exception ex52 = ex51;
					ProjectData.ClearProjectError();
				}
				try
				{
					this.A(A.PlotArea.Format.Line, B, C);
					return;
				}
				catch (Exception ex53)
				{
					ProjectData.SetProjectError(ex53);
					Exception ex54 = ex53;
					ProjectData.ClearProjectError();
					return;
				}
			}
			return;
		}
	}

	private void A(Chart A, Axis B, bool C, bool D, bool E, int F, int G)
	{
		Axis axis = B;
		if (axis.HasTitle)
		{
			AxisTitle axisTitle = axis.AxisTitle;
			if (C)
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
					this.A(axisTitle.Font, F, G);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
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
				try
				{
					this.A(axisTitle.Format.Fill, F, G);
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
			}
			if (E)
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
					this.A(axisTitle.Format.Line, F, G);
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
			}
			axisTitle = null;
		}
		if (C && axis.TickLabelPosition != XlTickLabelPosition.xlTickLabelPositionNone)
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
				this.A(axis.TickLabels.Font, F, G);
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
		}
		if (E)
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
			try
			{
				this.A(B, F, G);
			}
			catch (Exception ex9)
			{
				ProjectData.SetProjectError(ex9);
				Exception ex10 = ex9;
				ProjectData.ClearProjectError();
			}
		}
		axis = null;
		B = null;
	}

	private void A(TextRange2 A, int B, int C)
	{
		if (A.Font.Fill.ForeColor.RGB == B && !this.m_A.Contains(A.Font.Fill.ForeColor))
		{
			A.get_Characters(-1, -1).Font.Fill.ForeColor.RGB = C;
			this.m_A.Add(A.Font.Fill.ForeColor);
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.ChartFont A, int B, int C)
	{
		if (Operators.ConditionalCompareObjectEqual(A.Color, B, TextCompare: false) && !this.m_A.Contains(A))
		{
			A.Color = C;
			this.m_A.Add(A);
		}
	}

	private void A(Microsoft.Office.Core.ChartFont A, int B, int C)
	{
		if (Operators.ConditionalCompareObjectEqual(A.Color, B, TextCompare: false) && !this.m_A.Contains(A))
		{
			A.Color = C;
			this.m_A.Add(A);
		}
	}

	private void A(Microsoft.Office.Core.FillFormat A, int B, int C)
	{
		Microsoft.Office.Core.FillFormat fillFormat = A;
		if (fillFormat.ForeColor.RGB == 0)
		{
			if (fillFormat.ForeColor.ObjectThemeColor == MsoThemeColorIndex.msoThemeColorMixed)
			{
				goto IL_006a;
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
		}
		if (fillFormat.Visible == MsoTriState.msoTrue)
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
			this.A(fillFormat.ForeColor, B, C);
			this.A(fillFormat.BackColor, B, C);
		}
		goto IL_006a;
		IL_006a:
		fillFormat = null;
	}

	private void A(Microsoft.Office.Core.LineFormat A, int B, int C)
	{
		Microsoft.Office.Core.LineFormat lineFormat = A;
		if (lineFormat.ForeColor.RGB == 0)
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
			if (lineFormat.ForeColor.ObjectThemeColor == MsoThemeColorIndex.msoThemeColorMixed)
			{
				goto IL_0076;
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
		if (lineFormat.Visible == MsoTriState.msoTrue)
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
			this.A(lineFormat.ForeColor, B, C);
			this.A(lineFormat.BackColor, B, C);
		}
		goto IL_0076;
		IL_0076:
		lineFormat = null;
	}

	private void A(Microsoft.Office.Interop.PowerPoint.FillFormat A, int B, int C)
	{
		Microsoft.Office.Interop.PowerPoint.FillFormat fillFormat = A;
		if ((fillFormat.ForeColor.RGB != 0 || fillFormat.ForeColor.ObjectThemeColor != MsoThemeColorIndex.msoThemeColorMixed) && fillFormat.Visible == MsoTriState.msoTrue)
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
			this.A(fillFormat.ForeColor, B, C);
			this.A(fillFormat.BackColor, B, C);
		}
		fillFormat = null;
	}

	private void A(Microsoft.Office.Interop.PowerPoint.LineFormat A, int B, int C)
	{
		Microsoft.Office.Interop.PowerPoint.LineFormat lineFormat = A;
		if (lineFormat.ForeColor.RGB == 0)
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
			if (lineFormat.ForeColor.ObjectThemeColor == MsoThemeColorIndex.msoThemeColorMixed)
			{
				goto IL_006a;
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
		if (lineFormat.Visible == MsoTriState.msoTrue)
		{
			this.A(lineFormat.ForeColor, B, C);
			this.A(lineFormat.BackColor, B, C);
		}
		goto IL_006a;
		IL_006a:
		lineFormat = null;
	}

	private void A(Axis A, int B, int C)
	{
		Axis axis = A;
		if (axis.HasMajorGridlines)
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
			this.A(axis.MajorGridlines.Format.Line, B, C);
		}
		if (axis.HasMinorGridlines)
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
			this.A(axis.MinorGridlines.Format.Line, B, C);
		}
		try
		{
			Microsoft.Office.Interop.PowerPoint.LineFormat line = axis.Format.Line;
			if (line.Weight > 0f)
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
				if (line.Visible == MsoTriState.msoTrue)
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
					this.A(line.ForeColor, B, C);
				}
			}
			line = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		axis = null;
	}

	private void A(IMsoLegendKey A, int B, int C, bool D)
	{
		if (D)
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
					if (A.Format.Line.ForeColor.RGB == B)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								A.Format.Line.ForeColor.RGB = C;
								return;
							}
						}
					}
					return;
				}
			}
		}
		if (!Operators.ConditionalCompareObjectEqual(A.Interior.Color, B, TextCompare: false))
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
			A.Interior.Color = C;
			return;
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A, int B, int C, bool D, bool E, bool F)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.SmartArt.AllNodes.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				SmartArtNode smartArtNode = (SmartArtNode)enumerator.Current;
				if (D)
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
						Microsoft.Office.Core.TextFrame2 textFrame = smartArtNode.TextFrame2;
						if (textFrame.HasText == MsoTriState.msoTrue)
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
							TextRange2 textRange = textFrame.TextRange;
							int count = textRange.get_Runs(-1, -1).Count;
							for (int i = 1; i <= count; i = checked(i + 1))
							{
								this.B(textRange.get_Runs(i, -1).Font.Fill, B, C);
								this.C(textRange.get_Runs(i, -1).Font, B, C);
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
							textRange = null;
						}
						else
						{
							this.B(textFrame.TextRange.Font.Fill, B, C);
						}
						textFrame = null;
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
					enumerator2 = smartArtNode.Shapes.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Microsoft.Office.Core.Shape shape = (Microsoft.Office.Core.Shape)enumerator2.Current;
						if (E)
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
								this.B(shape.Fill, B, C);
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
						}
						if (!F)
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
						try
						{
							this.B(shape.Line, B, C);
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
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
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_01bd;
				}
				continue;
				end_IL_01bd:
				break;
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
		if (E)
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
			this.B(A, B, C);
		}
		if (F)
		{
			this.C(A, B, C);
		}
	}

	private void B(Microsoft.Office.Core.FillFormat A, int B, int C)
	{
		Microsoft.Office.Core.FillFormat fillFormat = A;
		if (fillFormat.Visible == MsoTriState.msoTrue)
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
			this.A(fillFormat.ForeColor, B, C);
			this.A(fillFormat.BackColor, B, C);
		}
		fillFormat = null;
	}

	private void B(Microsoft.Office.Core.LineFormat A, int B, int C)
	{
		Microsoft.Office.Core.LineFormat lineFormat = A;
		if (lineFormat.Visible == MsoTriState.msoTrue)
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
			this.A(lineFormat.ForeColor, B, C);
			this.A(lineFormat.BackColor, B, C);
		}
		lineFormat = null;
	}

	private void A(string A)
	{
		Forms.ErrorMessage(Window.GetWindow(this), A);
	}

	private void B(string A)
	{
		Forms.WarningMessage(Window.GetWindow(this), A);
	}

	private void C(string A)
	{
		Forms.InfoMessage(Window.GetWindow(this), A);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_B)
		{
			this.m_B = true;
			Uri resourceLocator = new Uri(AH.A(13196), UriKind.Relative);
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
				switch (2)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					radShape = (System.Windows.Controls.RadioButton)target;
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
					radSlide = (System.Windows.Controls.RadioButton)target;
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
					radPresentation = (System.Windows.Controls.RadioButton)target;
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
					chkFont = (System.Windows.Controls.CheckBox)target;
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
					chkFill = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkBorder = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkLayouts = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					scroller = (ScrollViewer)target;
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
					icColors = (ItemsControl)target;
					return;
				}
			}
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
					radRgb = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					radHex = (System.Windows.Controls.RadioButton)target;
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
					btnRefresh = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnRecolor = (System.Windows.Controls.Button)target;
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

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 10)
		{
			((System.Windows.Controls.Button)target).Click += ChooseColor;
		}
		if (connectionId == 11)
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
			((System.Windows.Controls.Button)target).Click += ChooseColor;
		}
		if (connectionId != 12)
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
			((System.Windows.Controls.Button)target).Click += ColorDequeued;
			return;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
