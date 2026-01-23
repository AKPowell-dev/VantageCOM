using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

[DesignerGenerated]
public sealed class wpfFootnoteSequence : System.Windows.Window, IComponentConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Range, int> A;

		public static Func<Range, int> B;

		public static Func<Range, int> C;

		public static Func<Range, int> D;

		public static Func<int, int> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal int A(Range A)
		{
			return A.Column;
		}

		[SpecialName]
		internal int B(Range A)
		{
			return A.Row;
		}

		[SpecialName]
		internal int C(Range A)
		{
			return A.Row;
		}

		[SpecialName]
		internal int D(Range A)
		{
			return A.Column;
		}

		[SpecialName]
		internal int A(int A)
		{
			return A;
		}
	}

	[CompilerGenerated]
	internal sealed class YF
	{
		public Match A;

		public YF(YF A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(int A)
		{
			return A > Conversions.ToInteger(this.A.ToString());
		}
	}

	private ObservableCollection<FootnoteError> m_A;

	[AccessedThroughProperty("grpInspect")]
	[CompilerGenerated]
	private GroupBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("radPrintAreas")]
	private RadioButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("radSelection")]
	private RadioButton B;

	[CompilerGenerated]
	[AccessedThroughProperty("grpSearch")]
	private GroupBox B;

	[AccessedThroughProperty("radRows")]
	[CompilerGenerated]
	private RadioButton C;

	[AccessedThroughProperty("radColumns")]
	[CompilerGenerated]
	private RadioButton D;

	[AccessedThroughProperty("lvErrors")]
	[CompilerGenerated]
	private ListView m_A;

	[AccessedThroughProperty("btnInspect")]
	[CompilerGenerated]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private Button B;

	private bool m_A;

	internal virtual GroupBox grpInspect
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

	internal virtual RadioButton radPrintAreas
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

	internal virtual RadioButton radSelection
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	internal virtual GroupBox grpSearch
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	internal virtual RadioButton radRows
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	internal virtual RadioButton radColumns
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			D = value;
		}
	}

	internal virtual ListView lvErrors
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

	internal virtual Button btnInspect
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
			RoutedEventHandler value2 = btnInspect_Click;
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
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual Button btnCancel
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public wpfFootnoteSequence()
	{
		base.Closing += wpfFootnoteSequence_Closing;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		lvErrors.SelectionChanged += lvErrors_SelectionChanged;
		if (K.Settings.FootnotesInspectPrintAreas)
		{
			radPrintAreas.IsChecked = true;
		}
		else
		{
			radSelection.IsChecked = true;
		}
		if (K.Settings.FootnotesSearchByRows)
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
			radRows.IsChecked = true;
		}
		else
		{
			radColumns.IsChecked = true;
		}
		btnInspect.Focus();
	}

	private void wpfFootnoteSequence_Closing(object sender, CancelEventArgs e)
	{
		if (base.DialogResult == true)
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
			e.Cancel = A();
		}
		if (!e.Cancel)
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
			this.m_A = null;
		}
		K.Settings.FootnotesInspectPrintAreas = radPrintAreas.IsChecked.Value;
		K.Settings.FootnotesSearchByRows = radRows.IsChecked.Value;
	}

	private void btnInspect_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = true;
	}

	private void lvErrors_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (lvErrors.SelectedIndex <= -1)
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
			Range range = this.m_A.ElementAt(lvErrors.SelectedIndex).Range;
			range.Worksheet.Activate();
			range.Select();
			_ = null;
			return;
		}
	}

	private bool A()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		List<Range> B = new List<Range>();
		List<string> C = new List<string>();
		if (radPrintAreas.IsChecked == true)
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
			{
				IEnumerator enumerator = application.ActiveWorkbook.Worksheets.GetEnumerator();
				try
				{
					while (enumerator.MoveNext())
					{
						Worksheet worksheet = (Worksheet)enumerator.Current;
						if (Operators.CompareString(worksheet.PageSetup.PrintArea, string.Empty, TextCompare: false) != 0)
						{
							string[] array = Strings.Split(worksheet.PageSetup.PrintArea, CultureInfo.CurrentCulture.TextInfo.ListSeparator, -1, CompareMethod.Text);
							foreach (string cell in array)
							{
								A(((_Worksheet)worksheet).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value)), ref B, ref C);
							}
						}
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_010c;
						}
						continue;
						end_IL_010c:
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
		else if (application.Selection is Range)
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
			A((Range)application.Selection, ref B, ref C);
		}
		int count = B.Count;
		checked
		{
			bool result;
			if (count == 0)
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
				result = false;
				Forms.InfoMessage(VH.A(146674));
			}
			else
			{
				result = true;
				this.m_A = new ObservableCollection<FootnoteError>();
				lvErrors.SelectionChanged -= lvErrors_SelectionChanged;
				int num = count - 1;
				for (int j = 0; j <= num; j++)
				{
					this.m_A.Add(new FootnoteError
					{
						SheetName = B[j].Worksheet.Name,
						Address = B[j].get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)),
						ErrorMessage = C[j],
						Range = B[j]
					});
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
				lvErrors.ItemsSource = this.m_A;
				lvErrors.Focus();
				lvErrors.SelectionChanged += lvErrors_SelectionChanged;
			}
			application = null;
			B = null;
			C = null;
			return result;
		}
	}

	private void A(Range A, ref List<Range> B, ref List<string> C)
	{
		Range range = null;
		Regex regex = Footnotes.FootnoteRegex();
		List<int> list = new List<int>();
		try
		{
			Range range2;
			if (Operators.ConditionalCompareObjectGreater(A.Cells.CountLarge, 1, TextCompare: false))
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
				range2 = A.SpecialCells(XlCellType.xlCellTypeConstants, RuntimeHelpers.GetObjectValue(Missing.Value));
				try
				{
					range = range2.SpecialCells(XlCellType.xlCellTypeVisible, RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				if (range != null)
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
					range2 = A.Application.Intersect(range2, range, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				}
			}
			else
			{
				range2 = A;
			}
			List<Range> list2 = new List<Range>();
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = range2.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range item = (Range)enumerator.Current;
					list2.Add(item);
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_020d;
					}
					continue;
					end_IL_020d:
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
			if (radRows.IsChecked == true)
			{
				List<Range> source = list2;
				Func<Range, int> keySelector;
				if (_Closure_0024__.A == null)
				{
					keySelector = (_Closure_0024__.A = [SpecialName] (Range range3) => range3.Column);
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
					keySelector = _Closure_0024__.A;
				}
				IOrderedEnumerable<Range> source2 = source.OrderBy(keySelector);
				Func<Range, int> keySelector2;
				if (_Closure_0024__.B == null)
				{
					keySelector2 = (_Closure_0024__.B = [SpecialName] (Range range3) => range3.Row);
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
					keySelector2 = _Closure_0024__.B;
				}
				list2 = source2.OrderBy(keySelector2).ToList();
			}
			else
			{
				IOrderedEnumerable<Range> source3 = list2.OrderBy([SpecialName] (Range range3) => range3.Row);
				Func<Range, int> keySelector3;
				if (_Closure_0024__.D == null)
				{
					keySelector3 = (_Closure_0024__.D = [SpecialName] (Range range3) => range3.Column);
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
					keySelector3 = _Closure_0024__.D;
				}
				list2 = source3.OrderBy(keySelector3).ToList();
			}
			using List<Range>.Enumerator enumerator2 = list2.GetEnumerator();
			IEnumerator enumerator3 = default(IEnumerator);
			IEnumerator enumerator4 = default(IEnumerator);
			YF yF = default(YF);
			while (enumerator2.MoveNext())
			{
				Range current = enumerator2.Current;
				if (!Footnotes.CellContainsFootnote(current, regex))
				{
					continue;
				}
				MatchCollection matchCollection = regex.Matches(current.Text.ToString());
				try
				{
					enumerator3 = matchCollection.GetEnumerator();
					while (enumerator3.MoveNext())
					{
						MatchCollection matchCollection2 = Regex.Matches(((Match)enumerator3.Current).Groups[1].ToString(), VH.A(146761));
						{
							enumerator4 = matchCollection2.GetEnumerator();
							try
							{
								while (enumerator4.MoveNext())
								{
									Match match = (Match)enumerator4.Current;
									list.Add(Conversions.ToInteger(match.ToString()));
									if (!list.Any())
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
									yF = new YF(yF);
									yF.A = match;
									if (list.Where(yF.A).Any())
									{
										B.Add(current);
										C.Add(VH.A(146768) + match.ToString() + VH.A(146789));
									}
									yF.A = null;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_0481;
									}
									continue;
									end_IL_0481:
									break;
								}
							}
							finally
							{
								IDisposable disposable = enumerator4 as IDisposable;
								if (disposable != null)
								{
									disposable.Dispose();
								}
							}
						}
						matchCollection2 = null;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_04b4;
						}
						continue;
						end_IL_04b4:
						break;
					}
				}
				finally
				{
					if (enumerator3 is IDisposable)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							(enumerator3 as IDisposable).Dispose();
							break;
						}
					}
				}
				matchCollection = null;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_04f0;
				}
				continue;
				end_IL_04f0:
				break;
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		checked
		{
			try
			{
				if (list.Any())
				{
					list = list.Distinct().ToList();
					List<int> source4 = list;
					Func<int, int> keySelector4;
					if (_Closure_0024__.A == null)
					{
						keySelector4 = (_Closure_0024__.A = [SpecialName] (int result) => result);
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
						keySelector4 = _Closure_0024__.A;
					}
					list = source4.OrderBy(keySelector4).ToList();
					int count = list.Count;
					if (count < list[count - 1])
					{
						int num = list[count - 1];
						for (int num2 = 1; num2 <= num; num2++)
						{
							if (list.Contains(num2))
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
							B.Add(A);
							C.Add(VH.A(146768) + num2 + VH.A(146832));
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_05fe;
							}
							continue;
							end_IL_05fe:
							break;
						}
					}
				}
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			Range range2 = null;
			list = null;
		}
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_A)
		{
			this.m_A = true;
			Uri resourceLocator = new Uri(VH.A(146857), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
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
					grpInspect = (GroupBox)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			radPrintAreas = (RadioButton)target;
			return;
		}
		if (connectionId == 3)
		{
			radSelection = (RadioButton)target;
			return;
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
					grpSearch = (GroupBox)target;
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
					radRows = (RadioButton)target;
					return;
				}
			}
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
					radColumns = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			lvErrors = (ListView)target;
			return;
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
					btnInspect = (Button)target;
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
					btnCancel = (Button)target;
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
}
