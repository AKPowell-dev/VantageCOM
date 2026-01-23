using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Xml;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Pagination;

namespace PowerPointAddIn1.Slides;

[DesignerGenerated]
public sealed class wpfSlideNums : Window, IComponentConnector, IStyleConnector
{
	private ObservableCollection<SlideNumber> m_A;

	private Microsoft.Office.Interop.PowerPoint.Application m_A;

	private bool m_A;

	private bool m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("lbxSlides")]
	private System.Windows.Controls.ListBox m_A;

	[AccessedThroughProperty("chkAll")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkStartAtOne")]
	private System.Windows.Controls.CheckBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSequential")]
	private System.Windows.Controls.CheckBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnApply")]
	private System.Windows.Controls.Button m_A;

	[AccessedThroughProperty("btnReset")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	[AccessedThroughProperty("btnClose")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_C;

	private bool m_C;

	internal virtual System.Windows.Controls.ListBox lbxSlides
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
			System.Windows.Input.KeyEventHandler value2 = lbxSlides_KeyUp;
			System.Windows.Controls.ListBox listBox = this.m_A;
			if (listBox != null)
			{
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
				switch (4)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				listBox.PreviewKeyDown += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkAll
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

	internal virtual System.Windows.Controls.Button btnApply
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
			RoutedEventHandler value2 = btnApply_Click;
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

	internal virtual System.Windows.Controls.Button btnReset
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
			RoutedEventHandler value2 = btnReset_Click;
			System.Windows.Controls.Button button = this.m_B;
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
			this.m_B = value;
			button = this.m_B;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnClose
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
			RoutedEventHandler value2 = btnClose_Click;
			System.Windows.Controls.Button button = this.m_C;
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

	public wpfSlideNums()
	{
		base.Loaded += wpfSlideNums_Loaded;
		base.Closing += wpfSlideNums_Closing;
		this.m_A = true;
		this.m_B = false;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void wpfSlideNums_Loaded(object sender, RoutedEventArgs e)
	{
		this.m_A = new ObservableCollection<SlideNumber>();
		this.m_A = NG.A.Application;
		List<int> list = new List<int>();
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = this.m_A.ActiveWindow.Selection.SlideRange.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Slide slide = (Slide)enumerator.Current;
					list.Add(slide.SlideID);
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		IEnumerator enumerator2 = default(IEnumerator);
		try
		{
			enumerator2 = this.m_A.ActivePresentation.Slides.GetEnumerator();
			IEnumerator enumerator3 = default(IEnumerator);
			while (enumerator2.MoveNext())
			{
				Slide slide2 = (Slide)enumerator2.Current;
				try
				{
					enumerator3 = slide2.CustomLayout.Shapes.GetEnumerator();
					while (true)
					{
						if (enumerator3.MoveNext())
						{
							if (!Numbers.IsSlideNumberPlaceholder((Microsoft.Office.Interop.PowerPoint.Shape)enumerator3.Current))
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
								string slideTitle = Helpers.GetSlideTitle(slide2);
								string tooltip;
								Bitmap bitmap;
								switch (Helpers.GetSlideType(slide2))
								{
								case SlideType.Title:
									tooltip = AH.A(119688);
									bitmap = OB.TitlePage;
									break;
								case SlideType.TableOfContents:
									tooltip = AH.A(119709);
									bitmap = OB.Numbering;
									break;
								case SlideType.Flysheet:
								case SlideType.Agenda:
									tooltip = AH.A(2597);
									bitmap = OB.Flysheet;
									break;
								case SlideType.Legal:
									tooltip = AH.A(119744);
									bitmap = OB.SetPertWeights;
									break;
								case SlideType.Contact:
									tooltip = AH.A(119771);
									bitmap = OB.Contact;
									break;
								case SlideType.Blank:
									tooltip = AH.A(119810);
									bitmap = OB.BlankSlide;
									break;
								case SlideType.CoverFront:
									tooltip = AH.A(119859);
									bitmap = OB.TitlePage;
									break;
								case SlideType.CoverBack:
									tooltip = AH.A(119882);
									bitmap = OB.TitlePage;
									break;
								default:
									tooltip = AH.A(119903);
									bitmap = OB.SlideLayoutGallery;
									break;
								}
								ObservableCollection<SlideNumber> a = this.m_A;
								SlideNumber obj = new SlideNumber
								{
									Slide = slide2,
									IsChecked = (slide2.HeadersFooters.SlideNumber.Visible == MsoTriState.msoTrue),
									IsSelected = list.Contains(slide2.SlideID),
									Image = Forms.GetImageSource(bitmap)
								};
								string text = AH.A(36272);
								string text2 = Conversions.ToString(slide2.SlideIndex);
								object obj2;
								if (slideTitle.Length != 0)
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
									obj2 = AH.A(119930) + slideTitle;
								}
								else
								{
									obj2 = "";
								}
								obj.Label = text + text2 + (string)obj2;
								obj.Tooltip = tooltip;
								a.Add(obj);
								break;
							}
							break;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_0303;
							}
							continue;
							end_IL_0303:
							break;
						}
						break;
					}
				}
				finally
				{
					if (enumerator3 is IDisposable)
					{
						while (true)
						{
							switch (7)
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
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_033d;
				}
				continue;
				end_IL_033d:
				break;
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
		lbxSlides.ItemsSource = this.m_A;
		lbxSlides.Focus();
		chkStartAtOne.IsChecked = KG.A.SlideNumbersStartAtOne;
		chkSequential.IsChecked = KG.A.SequentialSlideNumbers;
		A();
		B();
		chkAll.Checked += chkAll_CheckedChanged;
		chkAll.Unchecked += chkAll_CheckedChanged;
		chkStartAtOne.Checked += NumberSequenceChanged;
		chkStartAtOne.Unchecked += NumberSequenceChanged;
		chkSequential.Checked += NumberSequenceChanged;
		chkSequential.Unchecked += NumberSequenceChanged;
		lbxSlides.SelectionChanged += lbxSlides_SelectionChanged;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56735)).AddEventHandler(this.m_A, new EApplication_SlideSelectionChangedEventHandler(A));
		list = null;
	}

	private void wpfSlideNums_Closing(object sender, CancelEventArgs e)
	{
		C();
		XmlDocument settingsXml = KG.A.SettingsXml;
		try
		{
			XmlDocument settingsXml2 = KG.A.SettingsXml;
			settingsXml2.GetElementsByTagName(AH.A(119937)).Item(0).InnerText = chkStartAtOne.IsChecked.ToString();
			settingsXml2.GetElementsByTagName(AH.A(119970)).Item(0).InnerText = chkSequential.IsChecked.ToString();
			_ = null;
			KG.A.SaveSettings(settingsXml);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		settingsXml = null;
		chkAll.Checked -= chkAll_CheckedChanged;
		chkAll.Unchecked -= chkAll_CheckedChanged;
		chkStartAtOne.Checked -= NumberSequenceChanged;
		chkStartAtOne.Unchecked -= NumberSequenceChanged;
		chkSequential.Checked -= NumberSequenceChanged;
		chkSequential.Unchecked -= NumberSequenceChanged;
		lbxSlides.SelectionChanged -= lbxSlides_SelectionChanged;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56735)).RemoveEventHandler(this.m_A, new EApplication_SlideSelectionChangedEventHandler(A));
		this.m_A = null;
		JG.A(this.m_A);
		this.m_A = null;
	}

	private void lbxSlides_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Space || lbxSlides.SelectedItems.Count <= 0)
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
			bool isChecked = !((SlideNumber)lbxSlides.SelectedItems[0]).IsChecked;
			foreach (SlideNumber item in this.m_A)
			{
				if (!item.IsSelected)
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
				item.IsChecked = isChecked;
			}
			e.Handled = true;
			return;
		}
	}

	private void lbxSlides_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56735)).RemoveEventHandler(this.m_A, new EApplication_SlideSelectionChangedEventHandler(A));
		if (lbxSlides.SelectedItems.Count == 1)
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
			SlideNumber slideNumber = (SlideNumber)lbxSlides.SelectedItems[0];
			this.m_A = false;
			try
			{
				this.m_A.ActiveWindow.View.GotoSlide(slideNumber.Slide.SlideIndex);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				this.m_A.Remove(slideNumber);
				ProjectData.ClearProjectError();
			}
			slideNumber = null;
		}
		else if (lbxSlides.SelectedItems.Count > 1)
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
			this.m_A = false;
			List<int> list = new List<int>();
			foreach (SlideNumber item in this.m_A.Where([SpecialName] (SlideNumber A) => A.IsSelected).ToList())
			{
				list.Add(item.Slide.SlideIndex);
			}
			try
			{
				if (this.m_A.ActiveWindow.Panes[1].ViewType == PpViewType.ppViewThumbnails)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						this.m_A.ActiveWindow.Panes[1].Activate();
						this.m_A.ActivePresentation.Slides.Range(list.ToArray()).Select();
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
			list = null;
		}
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56735)).AddEventHandler(this.m_A, new EApplication_SlideSelectionChangedEventHandler(A));
	}

	private void A(SlideRange A)
	{
		if (!this.m_A)
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
					this.m_A = true;
					return;
				}
			}
		}
		List<int> list = new List<int>();
		IEnumerator enumerator = A.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Slide slide = (Slide)enumerator.Current;
				list.Add(slide.SlideID);
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_0056;
				}
				continue;
				end_IL_0056:
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
		System.Windows.Controls.ListBox listBox = lbxSlides;
		listBox.SelectionChanged -= lbxSlides_SelectionChanged;
		IEnumerator<SlideNumber> enumerator2 = default(IEnumerator<SlideNumber>);
		try
		{
			enumerator2 = this.m_A.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				SlideNumber current = enumerator2.Current;
				try
				{
					current.IsSelected = list.Contains(current.Slide.SlideID);
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
					break;
				default:
					goto end_IL_00e4;
				}
				continue;
				end_IL_00e4:
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
		listBox.SelectionChanged += lbxSlides_SelectionChanged;
		listBox = null;
		list = null;
	}

	private void SlideCheckedChanged(object sender, RoutedEventArgs e)
	{
		if (this.m_B)
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
			System.Windows.Controls.CheckBox checkBox = (System.Windows.Controls.CheckBox)sender;
			SlideNumber slideNumber = (SlideNumber)checkBox.DataContext;
			try
			{
				slideNumber.Slide.HeadersFooters.SlideNumber.Visible = ((checkBox.IsChecked == true) ? MsoTriState.msoTrue : MsoTriState.msoFalse);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				this.m_A.Remove(slideNumber);
				ProjectData.ClearProjectError();
			}
			checkBox = null;
			slideNumber = null;
			chkAll.Checked -= chkAll_CheckedChanged;
			chkAll.Unchecked -= chkAll_CheckedChanged;
			A();
			chkAll.Checked += chkAll_CheckedChanged;
			chkAll.Unchecked += chkAll_CheckedChanged;
			return;
		}
	}

	private void chkAll_CheckedChanged(object sender, RoutedEventArgs e)
	{
		int num;
		if (chkAll.IsChecked != true)
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
			num = 0;
		}
		else
		{
			num = -1;
		}
		MsoTriState visible = (MsoTriState)num;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = this.m_A.ActivePresentation.Slides.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Slide slide = (Slide)enumerator.Current;
				try
				{
					slide.HeadersFooters.SlideNumber.Visible = visible;
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
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_008d;
				}
				continue;
				end_IL_008d:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		this.m_B = true;
		bool value = chkAll.IsChecked.Value;
		IEnumerator<SlideNumber> enumerator2 = default(IEnumerator<SlideNumber>);
		try
		{
			enumerator2 = this.m_A.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				enumerator2.Current.IsChecked = value;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_00ff;
				}
				continue;
				end_IL_00ff:
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
		this.m_B = false;
		lbxSlides.Focus();
	}

	private void A()
	{
		System.Windows.Controls.CheckBox checkBox = chkAll;
		int num = this.m_A.Where([SpecialName] (SlideNumber A) => A.IsChecked).Count();
		if (num == 0)
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
			checkBox.IsChecked = false;
		}
		else if (num == lbxSlides.Items.Count)
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
			checkBox.IsChecked = true;
		}
		else
		{
			checkBox.IsChecked = null;
		}
		checkBox = null;
	}

	private void NumberSequenceChanged(object sender, RoutedEventArgs e)
	{
		B();
	}

	private void B()
	{
		System.Windows.Controls.Button button = btnApply;
		bool? isChecked;
		bool? flag = (isChecked = chkStartAtOne.IsChecked);
		bool? obj;
		if (flag.HasValue)
		{
			if (isChecked == true)
			{
				obj = true;
				goto IL_009b;
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
		}
		bool? isChecked2;
		flag = (isChecked2 = chkSequential.IsChecked);
		if (!flag.HasValue)
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
			obj = null;
		}
		else if (isChecked2 != true)
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
			obj = isChecked;
		}
		else
		{
			obj = true;
		}
		goto IL_009b;
		IL_009b:
		isChecked2 = obj;
		button.IsEnabled = isChecked2.Value;
	}

	private void btnApply_Click(object sender, RoutedEventArgs e)
	{
		if (Forms.OkCancelMessage(AH.A(120009)) == System.Windows.Forms.DialogResult.OK)
		{
			C();
			Numbers.Renumber();
		}
	}

	private void btnReset_Click(object sender, RoutedEventArgs e)
	{
		IEnumerator enumerator = this.m_A.ActivePresentation.Slides.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				SlideNumbers.Reset((Slide)enumerator.Current);
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

	private void btnClose_Click(object sender, RoutedEventArgs e)
	{
		Close();
	}

	private void C()
	{
		KG.A.SlideNumbersStartAtOne = chkStartAtOne.IsChecked.Value;
		KG.A.SequentialSlideNumbers = chkSequential.IsChecked.Value;
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
			switch (2)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_C = true;
			Uri resourceLocator = new Uri(AH.A(120349), UriKind.Relative);
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
					lbxSlides = (System.Windows.Controls.ListBox)target;
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
					chkAll = (System.Windows.Controls.CheckBox)target;
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
					chkStartAtOne = (System.Windows.Controls.CheckBox)target;
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
					chkSequential = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			btnApply = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnReset = (System.Windows.Controls.Button)target;
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
					btnClose = (System.Windows.Controls.Button)target;
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
	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId != 2)
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
			((System.Windows.Controls.CheckBox)target).Checked += SlideCheckedChanged;
			((System.Windows.Controls.CheckBox)target).Unchecked += SlideCheckedChanged;
			return;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
