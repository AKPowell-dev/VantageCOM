using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using A;
using MacabacusMacros;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.UI;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.UI;

[DesignerGenerated]
public sealed class wpfWarnings : UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<BaseError, bool> A;

		public static Func<BaseError, bool> B;

		public static Func<BaseError, int> A;

		public static Func<BaseError, BaseError> A;

		public static Func<int, IEnumerable<BaseError>, Y<int, IEnumerable<BaseError>>> A;

		public static Func<Y<int, IEnumerable<BaseError>>, int> A;

		public static Func<Y<int, IEnumerable<BaseError>>, Z<int, int>> A;

		public static Func<BaseError, bool> C;

		public static Func<BaseError, bool> D;

		public static Func<BaseError, bool> E;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(BaseError A)
		{
			return A.Slide == null;
		}

		[SpecialName]
		internal bool B(BaseError A)
		{
			return A.Slide != null;
		}

		[SpecialName]
		internal int A(BaseError A)
		{
			return A.Slide.SlideIndex;
		}

		[SpecialName]
		internal BaseError A(BaseError A)
		{
			return A;
		}

		[SpecialName]
		internal Y<int, IEnumerable<BaseError>> A(int A, IEnumerable<BaseError> B)
		{
			return new Y<int, IEnumerable<BaseError>>(A, B);
		}

		[SpecialName]
		internal int A(Y<int, IEnumerable<BaseError>> A)
		{
			return A.idx;
		}

		[SpecialName]
		internal Z<int, int> A(Y<int, IEnumerable<BaseError>> A)
		{
			return new Z<int, int>(A.idx, A.Group.Count());
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
	internal sealed class CD
	{
		public int A;

		public CD(CD A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(BaseError A)
		{
			if (A.Slide != null)
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
						return A.Slide.SlideIndex == this.A;
					}
				}
			}
			return false;
		}
	}

	[CompilerGenerated]
	internal sealed class DD
	{
		public Slide A;

		public DD(DD A)
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
		internal bool A(BaseError A)
		{
			return A.Slide.SlideIndex == this.A.SlideIndex;
		}
	}

	[CompilerGenerated]
	internal sealed class ED
	{
		public BaseError A;

		public Func<BaseError, bool> A;

		public Func<BaseError, bool> B;

		public Func<BaseError, bool> C;

		public ED(ED A)
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
		internal bool A(BaseError A)
		{
			return A.Shape == this.A.Shape;
		}

		[SpecialName]
		internal bool B(BaseError A)
		{
			return A.Shape == this.A.Shape;
		}

		[SpecialName]
		internal bool C(BaseError A)
		{
			return A.Shape == this.A.Shape;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private readonly double m_A;

	private List<BaseError> m_A;

	private DispatcherTimer m_A;

	private ICollectionView m_A;

	[CompilerGenerated]
	private ObservableCollection<BaseError> m_A;

	[CompilerGenerated]
	private List<BaseError> m_B;

	[CompilerGenerated]
	private int m_A;

	[CompilerGenerated]
	private bool m_A;

	[AccessedThroughProperty("lbxResults")]
	[CompilerGenerated]
	private ListBox m_A;

	private bool m_B;

	public ICollectionView SourceCollection
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

	private ObservableCollection<BaseError> AllItems
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

	private List<BaseError> ItemsQueuedToRemove
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

	internal virtual ListBox lbxResults
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
			KeyEventHandler value2 = lbxResults_KeyDown;
			ListBox listBox = this.m_A;
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
				switch (1)
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

	public wpfWarnings(List<BaseError> errors)
	{
		base.Loaded += wpfWarnings_Loaded;
		this.m_A = 700.0;
		this.m_A = null;
		NavigateAfterRemoveIndex = -1;
		InitializeComponent();
		this.m_A = errors;
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
			switch (2)
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

	private void wpfWarnings_Loaded(object sender, RoutedEventArgs e)
	{
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		List<BaseError> a = this.m_A;
		Func<BaseError, bool> predicate;
		if (_Closure_0024__.A == null)
		{
			predicate = (_Closure_0024__.A = [SpecialName] (BaseError A) => A.Slide == null);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			predicate = _Closure_0024__.A;
		}
		List<BaseError> list = a.Where(predicate).ToList();
		List<BaseError> a2 = this.m_A;
		Func<BaseError, bool> predicate2;
		if (_Closure_0024__.B == null)
		{
			predicate2 = (_Closure_0024__.B = [SpecialName] (BaseError A) => A.Slide != null);
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
			predicate2 = _Closure_0024__.B;
		}
		List<BaseError> list2 = a2.Where(predicate2).ToList();
		List<BaseError> source = list2;
		Func<BaseError, int> keySelector = [SpecialName] (BaseError A) => A.Slide.SlideIndex;
		Func<BaseError, BaseError> elementSelector;
		if (_Closure_0024__.A == null)
		{
			elementSelector = (_Closure_0024__.A = [SpecialName] (BaseError A) => A);
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
		IEnumerable<Y<int, IEnumerable<BaseError>>> source2 = source.GroupBy(keySelector, elementSelector, [SpecialName] (int A, IEnumerable<BaseError> B) => new Y<int, IEnumerable<BaseError>>(A, B));
		Func<Y<int, IEnumerable<BaseError>>, int> keySelector2;
		if (_Closure_0024__.A == null)
		{
			keySelector2 = (_Closure_0024__.A = [SpecialName] (Y<int, IEnumerable<BaseError>> A) => A.idx);
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
			keySelector2 = _Closure_0024__.A;
		}
		IOrderedEnumerable<Y<int, IEnumerable<BaseError>>> source3 = source2.OrderBy(keySelector2);
		Func<Y<int, IEnumerable<BaseError>>, Z<int, int>> selector;
		if (_Closure_0024__.A == null)
		{
			selector = (_Closure_0024__.A = [SpecialName] (Y<int, IEnumerable<BaseError>> A) => new Z<int, int>(A.idx, A.Group.Count()));
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
			selector = _Closure_0024__.A;
		}
		IEnumerable<Z<int, int>> enumerable = source3.Select(selector);
		using (IEnumerator<Z<int, int>> enumerator = enumerable.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				Z<int, int> current = enumerator.Current;
				dictionary.Add(current.SlideIndex, current.Count);
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_01b6;
				}
				continue;
				end_IL_01b6:
				break;
			}
		}
		AllItems = new ObservableCollection<BaseError>();
		if (list.Count > 0)
		{
			string text = Guid.NewGuid().ToString();
			int num = A(list);
			int num2 = B(list);
			int num3 = C(list);
			foreach (BaseError item in list)
			{
				((BaseError)item).SetVisualProperties(AH.A(58371), text, num, num2, num3);
				AllItems.Add(item);
			}
		}
		using (Dictionary<int, int>.Enumerator enumerator3 = dictionary.GetEnumerator())
		{
			CD cD = default(CD);
			while (enumerator3.MoveNext())
			{
				KeyValuePair<int, int> current3 = enumerator3.Current;
				cD = new CD(cD);
				cD.A = current3.Key;
				list2 = this.m_A.Where(cD.A).ToList();
				string text = Guid.NewGuid().ToString();
				int num = A(list2);
				int num2 = B(list2);
				int num3 = C(list2);
				string text2 = AH.A(36272) + current3.Key;
				using List<BaseError>.Enumerator enumerator4 = list2.GetEnumerator();
				while (enumerator4.MoveNext())
				{
					BaseError current4 = enumerator4.Current;
					((BaseError)current4).SetVisualProperties(text2, text, num, num2, num3);
					AllItems.Add(current4);
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_0370;
					}
					continue;
					end_IL_0370:
					break;
				}
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_0398;
				}
				continue;
				end_IL_0398:
				break;
			}
		}
		lbxResults.SelectionChanged -= ListBoxSelectionChanged;
		SourceCollection = CollectionViewSource.GetDefaultView(AllItems);
		SourceCollection.GroupDescriptions.Add(new PropertyGroupDescription(AH.A(52438)));
		SourceCollection.Filter = A;
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
			predicate = _Closure_0024__.E;
		}
		return A.Where(predicate).Count();
	}

	private void lbxResults_KeyDown(object sender, KeyEventArgs e)
	{
		Key key = e.Key;
		if (key <= Key.Space)
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
					if (key != Key.Escape)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								if (key != Key.Space || lbxResults.SelectedIndex <= -1)
								{
									return;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										C((BaseError)lbxResults.SelectedItem);
										return;
									}
								}
							}
						}
					}
					if (Callout.Dialog != null)
					{
						while (true)
						{
							switch (3)
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
					return;
				}
			}
		}
		if ((uint)(key - 23) > 3u)
		{
			while (true)
			{
				switch (4)
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
							switch (7)
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
				switch (3)
				{
				case 0:
					break;
				default:
					if (Callout.Dialog != null)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								Callout.Dialog.Top = -10000.0;
								Callout.A();
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

	private void NavKeyUp(object sender, KeyEventArgs e)
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
				switch (1)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (e.NewFocus is ToggleButton || Callout.DoNotClose)
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
							switch (1)
							{
							case 0:
								continue;
							}
							if (Callout.Dialog != null)
							{
								Callout.Dialog.Close();
							}
							ListBox listBox = lbxResults;
							if (listBox.SelectedIndex > -1)
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
								if (((BaseError)(BaseError)listBox.SelectedItem).IsFixed)
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
									ListBoxItem obj = e.OldFocus as ListBoxItem;
									bool? obj2;
									if (obj == null)
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
										obj2 = null;
									}
									else
									{
										obj2 = obj.IsKeyboardFocusWithin;
									}
									if (!object.Equals(obj2, true))
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
										RemoveItem((BaseError)listBox.SelectedItem, blnAnimate: true);
									}
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
				switch (3)
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
					if (Pane.ActiveItem != null)
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
						if (((BaseError)Pane.ActiveItem).IsFixed)
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
							RemoveItem(Pane.ActiveItem, blnAnimate: true);
						}
					}
					Pane.ActiveItem = baseError;
					if (!Keyboard.IsKeyDown(Key.Down))
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
						if (!Keyboard.IsKeyDown(Key.Up))
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
							if (!Keyboard.IsKeyDown(Key.Right) && !Keyboard.IsKeyDown(Key.Left))
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
								if (!Keyboard.IsKeyDown(Key.Next))
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
									if (!Keyboard.IsKeyDown(Key.Prior))
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
										A(baseError);
									}
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

	private void MarkAsRead(object obj, EventArgs ev)
	{
		((DispatcherTimer)obj).Stop();
		try
		{
			((BaseError)Pane.ActiveItem).FontWeight = FontWeights.Normal;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void A(BaseError A)
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56735)).RemoveEventHandler(NG.A.Application, new EApplication_SlideSelectionChangedEventHandler(Pane.A));
		try
		{
			int B = 0;
			if (Callout.Dialog != null)
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
				Callout.Dialog.A(A, ref B);
			}
			else
			{
				wpfCallout obj = new wpfCallout();
				obj.Top = -10000.0;
				obj.ShowActivated = false;
				obj.Show();
				obj.A(A, ref B);
				_ = null;
			}
			Pane.A(B);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			Forms.WarningMessage(AH.A(58396));
			ProjectData.ClearProjectError();
		}
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56735)).AddEventHandler(NG.A.Application, new EApplication_SlideSelectionChangedEventHandler(Pane.A));
	}

	private void B(BaseError A)
	{
		if (lbxResults.SelectedItem != A)
		{
			lbxResults.SelectedItem = A;
		}
	}

	public void ScrollSpy(Slide sld)
	{
		DD a = default(DD);
		DD CS_0024_003C_003E8__locals2 = new DD(a);
		CS_0024_003C_003E8__locals2.A = sld;
		if (!lbxResults.IsKeyboardFocusWithin)
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
			if (!lbxResults.IsKeyboardFocused)
			{
				try
				{
					List<BaseError> list = SourceCollection.OfType<BaseError>().ToList();
					BaseError baseError = list.FirstOrDefault([SpecialName] (BaseError A) => A.Slide.SlideIndex == CS_0024_003C_003E8__locals2.A.SlideIndex);
					if (baseError == null)
					{
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
					int index = list.IndexOf(baseError);
					ListBoxItem obj = (ListBoxItem)lbxResults.ItemContainerGenerator.ContainerFromItem(RuntimeHelpers.GetObjectValue(lbxResults.Items[index]));
					ScrollViewer scrollViewer = Forms.GetScrollViewer((DependencyObject)lbxResults) as ScrollViewer;
					Visual visual = obj;
					GroupItem groupItem = null;
					while (visual != null)
					{
						visual = VisualTreeHelper.GetParent(visual) as Visual;
						if (visual == null)
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
						Expander obj2 = (Expander)A(groupItem, typeof(Expander));
						ItemsPresenter relativeTo = (ItemsPresenter)scrollViewer.Content;
						scrollViewer.ScrollToVerticalOffset(obj2.TranslatePoint(default(Point), relativeTo).Y);
						relativeTo = null;
						Point point = default(Point);
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
		if (base.IsMouseOver || Callout.Dialog == null)
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
			Callout.Dialog.Close();
			return;
		}
	}

	private Visual A(Visual A, Type B)
	{
		if (A == null)
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
					return null;
				}
			}
		}
		if (A.GetType() == B)
		{
			while (true)
			{
				switch (7)
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
				switch (6)
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
			int num2 = 0;
			while (true)
			{
				if (num2 <= num)
				{
					Visual a = VisualTreeHelper.GetChild(A, num2) as Visual;
					visual = this.A(a, B);
					if (visual != null)
					{
						break;
					}
					while (true)
					{
						switch (6)
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
					num2++;
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
				break;
			}
			return visual;
		}
	}

	private void FixButtonClicked(object sender, RoutedEventArgs e)
	{
		C((BaseError)((Button)sender).DataContext);
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
		Fixes.ShowOptions(baseError, (ToggleButton)sender, blnRefocusPane: true);
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
		A();
		CollapseAnimationRunning = false;
	}

	private void A()
	{
		B();
		AllItems.Remove(ItemsQueuedToRemove[0]);
		Main.Analysis.Errors.Remove(ItemsQueuedToRemove[0]);
		ItemsQueuedToRemove.RemoveAt(0);
		if (NavigateAfterRemoveIndex <= -1)
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
			lbxResults.SelectedIndex = NavigateAfterRemoveIndex;
			Pane.A(lbxResults);
			NavigateAfterRemoveIndex = -1;
			return;
		}
	}

	private void B()
	{
		//IL_0051: Unknown result type (might be due to invalid IL or missing references)
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		//IL_0059: Invalid comparison between Unknown and I4
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		//IL_007b: Unknown result type (might be due to invalid IL or missing references)
		//IL_007e: Invalid comparison between Unknown and I4
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
				switch (1)
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
		foreach (BaseError item in list)
		{
			((BaseError)item).ErrorsCount = num;
			((BaseError)item).WarningsCount = num2;
			((BaseError)item).MessagesCount = num3;
		}
		list = null;
	}

	public void ToggleItems()
	{
		ApplyFilter();
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Pane.A(lbxResults);
			return;
		}
	}

	public void ApplyFilter()
	{
		lbxResults.SelectionChanged -= ListBoxSelectionChanged;
		Pane.CloseCallout();
		Pane.ActiveItem = null;
		SourceCollection = CollectionViewSource.GetDefaultView(lbxResults.ItemsSource);
		SourceCollection.Filter = A;
		SourceCollection.Refresh();
		lbxResults.SelectionChanged += ListBoxSelectionChanged;
		lbxResults.Focus();
	}

	private bool A(object A)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Invalid comparison between Unknown and I4
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0064: Invalid comparison between Unknown and I4
		//IL_009c: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a4: Invalid comparison between Unknown and I4
		BaseError baseError = (BaseError)A;
		if ((int)((BaseError)baseError).Severity == 3)
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
			if (Pane.TaskPane.chkErrors.IsChecked == true)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						return this.A(baseError);
					}
				}
			}
		}
		if ((int)((BaseError)baseError).Severity == 2 && Pane.TaskPane.chkWarnings.IsChecked == true)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					return this.A(baseError);
				}
			}
		}
		if ((int)((BaseError)baseError).Severity == 1)
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
			if (Pane.TaskPane.chkMessages.IsChecked == true)
			{
				while (true)
				{
					switch (3)
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
		string text = Pane.TaskPane.txtSearch.Text.ToLower();
		if (text.Length == 0)
		{
			return true;
		}
		if (!((BaseError)A).Title.ToLower().Contains(text))
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
			if (!((BaseError)A).Subtitle.ToLower().Contains(text))
			{
				return false;
			}
		}
		return true;
	}

	private void RemoveButtonCicked(object sender, RoutedEventArgs e)
	{
		RemoveItemAndNavigate((BaseError)((Button)sender).DataContext);
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
				ED a = default(ED);
				ED CS_0024_003C_003E8__locals22 = new ED(a);
				CS_0024_003C_003E8__locals22.A = itm;
				try
				{
					if (CS_0024_003C_003E8__locals22.A is AbbreviationSpacing)
					{
						IEnumerator<BaseError> enumerator = default(IEnumerator<BaseError>);
						IEnumerator<TextRange2> enumerator3 = default(IEnumerator<TextRange2>);
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
							try
							{
								List<BaseError> errors = Main.Analysis.Errors;
								Func<BaseError, bool> predicate;
								if (CS_0024_003C_003E8__locals22.A != null)
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
									predicate = CS_0024_003C_003E8__locals22.A;
								}
								else
								{
									predicate = (CS_0024_003C_003E8__locals22.A = [SpecialName] (BaseError A) => A.Shape == CS_0024_003C_003E8__locals22.A.Shape);
								}
								enumerator = errors.Where(predicate).GetEnumerator();
								while (enumerator.MoveNext())
								{
									BaseError current = enumerator.Current;
									if (!(current is AbbreviationMillions))
									{
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
									foreach (TextRange2 textRange2 in ((BaseError)CS_0024_003C_003E8__locals22.A).TextRanges)
									{
										try
										{
											enumerator3 = ((BaseError)current).TextRanges.GetEnumerator();
											while (true)
											{
												if (enumerator3.MoveNext())
												{
													TextRange2 current3 = enumerator3.Current;
													if (current3 == null)
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
													if (current3.Start < textRange2.Start || current3.Start >= textRange2.Start + textRange2.Length)
													{
														continue;
													}
													List<string> list = new List<string>(new string[12]
													{
														AH.A(15034),
														AH.A(15010),
														AH.A(15024),
														AH.A(15029),
														AH.A(8238),
														AH.A(8040),
														AH.A(15000),
														AH.A(15005),
														AH.A(8136),
														AH.A(7938),
														AH.A(8103),
														AH.A(7905)
													});
													Match match = new Regex(AH.A(17795) + Strings.Join(list.ToArray(), AH.A(58688)) + AH.A(14255), RegexOptions.None).Match(textRange2.Text);
													if (match != null && match.Success)
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
														Match match2 = match;
														((BaseError)current).TextRanges[((BaseError)current).TextRanges.IndexOf(current3)] = textRange2.get_Characters(match2.Index + 1, match2.Length);
														match2 = null;
														match = null;
													}
													list = null;
													break;
												}
												while (true)
												{
													switch (6)
													{
													case 0:
														break;
													default:
														goto end_IL_02b1;
													}
													continue;
													end_IL_02b1:
													break;
												}
												break;
											}
										}
										finally
										{
											if (enumerator3 != null)
											{
												while (true)
												{
													switch (1)
													{
													case 0:
														continue;
													}
													enumerator3.Dispose();
													break;
												}
											}
										}
									}
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_02fc;
									}
									continue;
									end_IL_02fc:
									break;
								}
							}
							finally
							{
								enumerator?.Dispose();
							}
							break;
						}
					}
					else
					{
						if (CS_0024_003C_003E8__locals22.A is AbbreviationMillions)
						{
							goto IL_033d;
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
						if (CS_0024_003C_003E8__locals22.A is AbbreviationBillions)
						{
							goto IL_033d;
						}
						if (CS_0024_003C_003E8__locals22.A is IeEgComma)
						{
							Regex regex = new Regex(AH.A(58717), RegexOptions.IgnoreCase);
							IEnumerator<BaseError> enumerator4 = default(IEnumerator<BaseError>);
							try
							{
								List<BaseError> errors2 = Main.Analysis.Errors;
								Func<BaseError, bool> predicate2;
								if (CS_0024_003C_003E8__locals22.C != null)
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
									predicate2 = CS_0024_003C_003E8__locals22.C;
								}
								else
								{
									predicate2 = (CS_0024_003C_003E8__locals22.C = [SpecialName] (BaseError A) => A.Shape == CS_0024_003C_003E8__locals22.A.Shape);
								}
								enumerator4 = errors2.Where(predicate2).GetEnumerator();
								IEnumerator<TextRange2> enumerator5 = default(IEnumerator<TextRange2>);
								IEnumerator<TextRange2> enumerator6 = default(IEnumerator<TextRange2>);
								while (enumerator4.MoveNext())
								{
									BaseError current4 = enumerator4.Current;
									if (!(current4 is GrammarIeEg))
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
										enumerator5 = ((BaseError)CS_0024_003C_003E8__locals22.A).TextRanges.GetEnumerator();
										while (enumerator5.MoveNext())
										{
											TextRange2 current5 = enumerator5.Current;
											try
											{
												enumerator6 = ((BaseError)current4).TextRanges.GetEnumerator();
												while (enumerator6.MoveNext())
												{
													TextRange2 current6 = enumerator6.Current;
													if (current6 == null)
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
													if (current6.Start < current5.Start || current6.Start >= current5.Start + current5.Length)
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
														Match match3 = regex.Match(current5.Text);
														if (match3 == null)
														{
															break;
														}
														while (true)
														{
															switch (5)
															{
															case 0:
																continue;
															}
															if (!match3.Success)
															{
																break;
															}
															while (true)
															{
																switch (4)
																{
																case 0:
																	continue;
																}
																Match match4 = match3;
																((BaseError)current4).TextRanges[((BaseError)current4).TextRanges.IndexOf(current6)] = current5.get_Characters(match4.Index + 1, match4.Length);
																match4 = null;
																match3 = null;
																break;
															}
															break;
														}
														break;
													}
													break;
												}
											}
											finally
											{
												if (enumerator6 != null)
												{
													while (true)
													{
														switch (4)
														{
														case 0:
															continue;
														}
														enumerator6.Dispose();
														break;
													}
												}
											}
										}
										while (true)
										{
											switch (1)
											{
											case 0:
												break;
											default:
												goto end_IL_0798;
											}
											continue;
											end_IL_0798:
											break;
										}
									}
									finally
									{
										if (enumerator5 != null)
										{
											while (true)
											{
												switch (3)
												{
												case 0:
													continue;
												}
												enumerator5.Dispose();
												break;
											}
										}
									}
								}
								while (true)
								{
									switch (4)
									{
									case 0:
										break;
									default:
										goto end_IL_07c6;
									}
									continue;
									end_IL_07c6:
									break;
								}
							}
							finally
							{
								if (enumerator4 != null)
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											continue;
										}
										enumerator4.Dispose();
										break;
									}
								}
							}
							regex = null;
						}
					}
					goto end_IL_001b;
					IL_033d:
					IEnumerator<BaseError> enumerator7 = default(IEnumerator<BaseError>);
					try
					{
						List<BaseError> errors3 = Main.Analysis.Errors;
						Func<BaseError, bool> predicate3;
						if (CS_0024_003C_003E8__locals22.B != null)
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
							predicate3 = CS_0024_003C_003E8__locals22.B;
						}
						else
						{
							predicate3 = (CS_0024_003C_003E8__locals22.B = [SpecialName] (BaseError A) => A.Shape == CS_0024_003C_003E8__locals22.A.Shape);
						}
						enumerator7 = errors3.Where(predicate3).GetEnumerator();
						while (enumerator7.MoveNext())
						{
							BaseError current7 = enumerator7.Current;
							if (!(current7 is AbbreviationSpacing) || ((BaseError)current7).TextRanges[0] == null)
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
							TextRange2 textRange = ((BaseError)CS_0024_003C_003E8__locals22.A).TextRanges[0];
							if (textRange.Start >= ((BaseError)current7).TextRanges[0].Start)
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
								if (textRange.Start < ((BaseError)current7).TextRanges[0].Start + ((BaseError)current7).TextRanges[0].Length)
								{
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										List<string> list2 = new List<string>(new string[12]
										{
											AH.A(15034),
											AH.A(15010),
											AH.A(15024),
											AH.A(15029),
											AH.A(8238),
											AH.A(8040),
											AH.A(15000),
											AH.A(15005),
											AH.A(8136),
											AH.A(7938),
											AH.A(8103),
											AH.A(7905)
										});
										((BaseError)current7).ReplacementText[0] = Regex.Replace(((BaseError)current7).ReplacementText[0], AH.A(58691) + Strings.Join(list2.ToArray(), AH.A(58688)) + AH.A(14255), AH.A(58712) + ((BaseError)CS_0024_003C_003E8__locals22.A).ReplacementText[0]);
										list2 = null;
										break;
									}
									break;
								}
							}
							textRange = null;
						}
					}
					finally
					{
						if (enumerator7 != null)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								enumerator7.Dispose();
								break;
							}
						}
					}
					end_IL_001b:;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				CS_0024_003C_003E8__locals22.A = null;
			}
			if (ItemsQueuedToRemove == null)
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
				ItemsQueuedToRemove = new List<BaseError>();
			}
			ItemsQueuedToRemove.Add(itm);
			if (blnAnimate)
			{
				while (true)
				{
					switch (7)
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
								A();
							}
						}
						catch (NullReferenceException ex3)
						{
							ProjectData.SetProjectError(ex3);
							NullReferenceException ex4 = ex3;
							clsReporting.LogException((Exception)ex4);
							ProjectData.ClearProjectError();
						}
						listBoxItem = null;
						return;
					}
					}
				}
			}
			A();
		}
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
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_B = true;
			Uri resourceLocator = new Uri(AH.A(58748), UriKind.Relative);
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
		if (connectionId == 1)
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
					lbxResults = (ListBox)target;
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

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 2)
		{
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = UIElement.LostKeyboardFocusEvent;
			eventSetter.Handler = new KeyboardFocusChangedEventHandler(ListBoxItemLostKeyboardFocus);
			((Style)target).Setters.Add(eventSetter);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			((Button)target).Click += RemoveButtonCicked;
		}
		if (connectionId == 4)
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
			((Button)target).Click += FixButtonClicked;
		}
		if (connectionId == 5)
		{
			((ToggleButton)target).Checked += ShowFixOptions;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}

	[SpecialName]
	[CompilerGenerated]
	private bool B(BaseError A)
	{
		return Operators.CompareString(((BaseError)A).Guid, ((BaseError)ItemsQueuedToRemove[0]).Guid, TextCompare: false) == 0;
	}
}
