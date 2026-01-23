using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Macabacus_Word.TextOps.Redaction.Redactors;
using Macabacus_Word.TextOps.Redaction.Values;
using Macabacus_Word.Values;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.TextOps.Redaction.Process;

[DesignerGenerated]
public sealed class wpfRedactSelection : System.Windows.Window, IComponentConnector
{
	[CompilerGenerated]
	internal sealed class RB
	{
		public Exception A;

		public wpfRedactSelection A;

		[SpecialName]
		internal void A()
		{
			Forms.ErrorMessage(System.Windows.Window.GetWindow(this.A), XC.A(44220) + this.A.Message);
		}
	}

	private SelectionValue m_A;

	private Selection m_A;

	private BackgroundWorker m_A;

	private int m_A;

	private object m_A;

	private object m_B;

	private bool m_A;

	private RangeValue m_A;

	[CompilerGenerated]
	private bool m_B;

	private readonly object m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private Button m_A;

	[AccessedThroughProperty("pbRedact")]
	[CompilerGenerated]
	private ProgressBar m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("txtRedact")]
	private TextBlock m_A;

	private bool m_C;

	public bool ProgressBarLaunched
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

	internal virtual Button btnCancel
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
			MouseEventHandler value2 = btnCancel_MouseEnter;
			MouseEventHandler value3 = btnCancel_MouseLeave;
			RoutedEventHandler value4 = btnCancel_Click;
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
				button.MouseEnter -= value2;
				button.MouseLeave -= value3;
				button.Click -= value4;
			}
			this.m_A = value;
			button = this.m_A;
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
				button.MouseEnter += value2;
				button.MouseLeave += value3;
				button.Click += value4;
				return;
			}
		}
	}

	internal virtual ProgressBar pbRedact
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

	internal virtual TextBlock txtRedact
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

	public wpfRedactSelection()
	{
		base.Loaded += wpfRedactRefresh_Loaded;
		base.Closing += wpfRedact_Closing;
		base.MouseDown += Window_MouseDown;
		this.m_A = 1;
		this.m_A = false;
		ProgressBarLaunched = false;
		this.m_C = XC.A(18757);
		InitializeComponent();
		this.m_A = new SelectionValue(PC.A.Application);
		this.m_A = this.m_A.WdApp.Selection;
		SmartArtRedactor.WpfRedactSelection = this;
		SmartArtRedactor.ApprovedMultipleBulletsInSmartArt = false;
	}

	private void wpfRedactRefresh_Loaded(object sender, RoutedEventArgs e)
	{
		this.m_A = new BackgroundWorker();
		BackgroundWorker a = this.m_A;
		a.WorkerSupportsCancellation = true;
		a.WorkerReportsProgress = true;
		a.DoWork += bgw_DoWork;
		a.ProgressChanged += bgw_ProgressChanged;
		a.RunWorkerCompleted += bgw_Completed;
		_ = null;
		txtRedact.Text = XC.A(18464);
		pbRedact.IsIndeterminate = true;
		this.m_A.RunWorkerAsync();
		ProgressBarLaunched = true;
	}

	private void bgw_DoWork(object sender, DoWorkEventArgs e)
	{
		A(e);
	}

	public void RedactNoProgressBar()
	{
		if (this.m_A != null)
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
			A((DoWorkEventArgs)null);
			return;
		}
	}

	private void A(DoWorkEventArgs A)
	{
		UndoRecord undo = RedactUtilities.BeginRedactionProcess(XC.A(980), this.m_A.WdApp);
		try
		{
			RedactUtilities.ExpandAtInsertionPoint(this.m_A);
			this.m_A = this.m_A.Range.Paragraphs.Count;
			if (!this.A(A))
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
					if (!A.Cancel)
					{
						I();
						Thread.Sleep(300);
					}
					return;
				}
			}
		}
		catch (MultipleBulletsInSmartArtException ex)
		{
			ProjectData.SetProjectError(ex);
			MultipleBulletsInSmartArtException ex2 = ex;
			ProjectData.ClearProjectError();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Exception A2 = ex4;
			base.Dispatcher.Invoke([SpecialName] () =>
			{
				Forms.ErrorMessage(System.Windows.Window.GetWindow(this), XC.A(44220) + A2.Message);
			});
			clsReporting.LogException(A2);
			ProjectData.ClearProjectError();
		}
		finally
		{
			RedactUtilities.CompleteRedactionProcess(XC.A(18529), isFindAndRedact: false, ref undo, this.m_A);
			undo = null;
		}
	}

	private bool A(DoWorkEventArgs A)
	{
		H();
		if (clsUtilities.IsSelectionTypeOfShape(this.m_A))
		{
			if (!RedactSelectedShapes(this.m_A, A))
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
						return false;
					}
				}
			}
		}
		else if (clsUtilities.IsSelectionTextInsideTable(this.m_A))
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
			RedactSelectedTextInsideTable(this.m_A, A);
		}
		else if (this.m_A.Type == WdSelectionType.wdSelectionNormal)
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
			if (!RedactNormalSelection(this.m_A, A))
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						return false;
					}
				}
			}
		}
		return true;
	}

	private bool RedactSelectedShapes(Selection sel, DoWorkEventArgs e)
	{
		if (B(e))
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
					return false;
				}
			}
		}
		this.m_A = new RangeValue(sel.Range, this.m_A);
		if (!D())
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					return false;
				}
			}
		}
		if (!C())
		{
			return false;
		}
		E();
		ref object a = ref this.m_A;
		a = Operators.AddObject(a, 1);
		H();
		return true;
	}

	private void RedactSelectedTextInsideTable(Selection sel, DoWorkEventArgs e)
	{
		IEnumerator enumerator = default(IEnumerator);
		if (sel.Cells.Count > 1)
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
					try
					{
						enumerator = sel.Cells.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Cell cell = (Cell)enumerator.Current;
							if (B(e))
							{
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
							this.m_A = new RangeValue(cell.Range, this.m_A);
							E();
							ref object a = ref this.m_A;
							a = Operators.AddObject(a, 1);
							H();
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
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									(enumerator as IDisposable).Dispose();
									goto end_IL_00b5;
								}
								continue;
								end_IL_00b5:
								break;
							}
						}
					}
				}
			}
		}
		if (B(e))
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
			this.m_A = new RangeValue(sel.Range, this.m_A);
			E();
			ref object a2 = ref this.m_A;
			a2 = Operators.AddObject(a2, 1);
			H();
			return;
		}
	}

	private bool RedactNormalSelection(Selection sel, DoWorkEventArgs e)
	{
		List<Range> list = A(sel);
		using (List<Range>.Enumerator enumerator = list.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				Range current = enumerator.Current;
				if (B(e))
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
							return false;
						}
					}
				}
				this.m_A = new RangeValue(current, this.m_A);
				if (!D())
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							return false;
						}
					}
				}
				E();
				ref object a = ref this.m_A;
				a = Operators.AddObject(a, 1);
				H();
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_00a1;
				}
				continue;
				end_IL_00a1:
				break;
			}
		}
		return true;
	}

	private void A()
	{
		if (this.m_A.HasFloatingShapes)
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
					RedactUtilities.RedactTextInShape(this.m_A.FloatingShapesList.First(), this.m_A.Range);
					return;
				}
			}
		}
		if (this.m_A.HasGroupedFloatingShapes)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					RedactUtilities.RedactTextInShape(this.m_A.GroupedFloatingShapesList.First().FloatingShapesList.First(), this.m_A.Range);
					return;
				}
			}
		}
		if (this.m_A.HasGroupedNonFloatingShapes)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					RedactUtilities.RedactTextInShape(this.m_A.GroupedNonFloatingShapesList.First().FloatingShapesList.First(), this.m_A.Range);
					return;
				}
			}
		}
		RedactUtilities.RedactTextInShape(this.m_A.NonFloatingShapesList.First(), this.m_A.Range);
	}

	private void B()
	{
		RedactUtilities.RedactFloatingShapes(this.m_A.FloatingShapesList);
		A(this.m_A.GroupedFloatingShapesList);
	}

	private void C()
	{
		if (this.m_A.HasNonFloatingShapes)
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
					RedactUtilities.RedactInlineShape(this.m_A.NonFloatingShapesList.First());
					return;
				}
			}
		}
		RedactUtilities.RedactFloatingShapes(this.m_A.FloatingShapesList);
		A(this.m_A.GroupedNonFloatingShapesList);
	}

	private void D()
	{
		if (A())
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
					F();
					return;
				}
			}
		}
		G();
	}

	private void E()
	{
		if (this.m_A.IsTextInsideShape)
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
					A();
					return;
				}
			}
		}
		if (this.m_A.IsNonFloatingShape)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					C();
					return;
				}
			}
		}
		if (this.m_A.IsFloatingShape)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					B();
					return;
				}
			}
		}
		if (this.m_A.SelType != WdSelectionType.wdSelectionNormal)
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
			if (!this.m_A.IsTextInsideTable)
			{
				return;
			}
		}
		D();
	}

	private void F()
	{
		Range range = this.m_A.Range;
		List<IShape> floatingShapesList = this.m_A.FloatingShapesList;
		List<IShape> nonFloatingShapesList = this.m_A.NonFloatingShapesList;
		List<GroupedShapesValue> groupedFloatingShapesList = this.m_A.GroupedFloatingShapesList;
		List<GroupedShapesValue> groupedNonFloatingShapesList = this.m_A.GroupedNonFloatingShapesList;
		if (B())
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
			range = A(range);
		}
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
			if (RedactUtilities.DoesRangeContainParagraphMark(range))
			{
				RedactUtilities.RedactFloatingShapes(floatingShapesList);
				A(groupedFloatingShapesList);
			}
		}
		if (this.m_A.HasNonFloatingShapes)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					A(range, nonFloatingShapesList);
					return;
				}
			}
		}
		if (this.m_A.HasGroupedNonFloatingShapes)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					A(groupedNonFloatingShapesList);
					return;
				}
			}
		}
		RedactUtilities.RedactText(range, redactEntireLine: true);
	}

	private void G()
	{
		Range range = this.m_A.Range;
		if (!clsUtilities.IsRangeInHeaderFooter(range))
		{
			RedactUtilities.RedactText(range, redactEntireLine: true);
		}
		else
		{
			RedactUtilities.RedactText(range, redactEntireLine: false);
		}
		range = null;
	}

	private Range A(Range A)
	{
		Range range = B(A);
		if (range == null)
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
					return A;
				}
			}
		}
		return this.A(A, range);
	}

	private void A(List<GroupedShapesValue> A)
	{
		using List<GroupedShapesValue>.Enumerator enumerator = A.GetEnumerator();
		while (enumerator.MoveNext())
		{
			RedactUtilities.RedactFloatingShapes(enumerator.Current.FloatingShapesList);
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

	private void A(Range A, List<IShape> B)
	{
		foreach (IShape item in B)
		{
			RedactUtilities.RedactInlineShape(item);
		}
		if (A.Start < B[0].RangeStart())
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
			Document activeDocument = this.m_A.WdApp.ActiveDocument;
			Range range;
			object Start = (range = A).Start;
			object End = B[0].RangeStart();
			Range range2 = activeDocument.Range(ref Start, ref End);
			range.Start = Conversions.ToInteger(Start);
			RedactUtilities.RedactText(range2.Duplicate, redactEntireLine: true);
		}
		checked
		{
			int num = B[0].RangeEnd() + 1;
			int num2 = B.Count - 1;
			for (int i = 1; i <= num2; i++)
			{
				if (num < B[i].RangeStart())
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
					Document activeDocument2 = this.m_A.WdApp.ActiveDocument;
					object End = num;
					object Start = B[i].RangeStart();
					Range range3 = activeDocument2.Range(ref End, ref Start);
					num = Conversions.ToInteger(End);
					RedactUtilities.RedactText(range3.Duplicate, redactEntireLine: true);
				}
				num = B[i].RangeEnd() + 1;
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				if (num >= A.End)
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
					Document activeDocument3 = this.m_A.WdApp.ActiveDocument;
					object Start = num;
					Range range;
					object End = (range = A).End;
					Range range4 = activeDocument3.Range(ref Start, ref End);
					range.End = Conversions.ToInteger(End);
					num = Conversions.ToInteger(Start);
					RedactUtilities.RedactText(range4.Duplicate, redactEntireLine: true);
					return;
				}
			}
		}
	}

	private List<Range> A(Selection A)
	{
		List<Range> list = new List<Range>();
		Range range = A.Range;
		int count = range.Paragraphs.Count;
		for (int i = 1; i <= count; i = checked(i + 1))
		{
			Range range2 = range.Paragraphs[i].Range;
			if (i == 1)
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
				range2.Start = range.Start;
			}
			if (i == range.Paragraphs.Count)
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
				range2.End = range.End;
			}
			list.Add(range2);
		}
		return list;
	}

	private Range B(Range A)
	{
		Range firstWord = TextRedactor.GetFirstWord(A.Duplicate, this.m_A.WdApp);
		if (firstWord == null)
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
					return null;
				}
			}
		}
		int start = firstWord.Start;
		int end = firstWord.End;
		TextRedactor.RedactFirstOccurrence(firstWord, firstWord.Text, this.m_A.WdApp.ActiveDocument);
		Document activeDocument = this.m_A.WdApp.ActiveDocument;
		object Start = start;
		object End = end;
		Range result = activeDocument.Range(ref Start, ref End);
		end = Conversions.ToInteger(End);
		start = Conversions.ToInteger(Start);
		return result;
	}

	private Range A(Range A, Range B)
	{
		Document activeDocument = this.m_A.WdApp.ActiveDocument;
		Range range;
		object Start = (range = B).End;
		Range range2;
		object End = (range2 = A).End;
		Range result = activeDocument.Range(ref Start, ref End);
		range2.End = Conversions.ToInteger(End);
		range.End = Conversions.ToInteger(Start);
		return result;
	}

	private bool A()
	{
		if (!this.m_A.HasNonFloatingShapes)
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
			if (!this.m_A.HasFloatingShapes)
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
				if (!this.m_A.HasGroupedFloatingShapes)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							return this.m_A.HasGroupedNonFloatingShapes;
						}
					}
				}
			}
		}
		return true;
	}

	private bool B()
	{
		if (!this.m_A.HasFloatingShapes)
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
					return this.m_A.HasGroupedFloatingShapes;
				}
			}
		}
		return true;
	}

	private bool C()
	{
		if (this.m_A)
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
					return true;
				}
			}
		}
		using (List<IShape>.Enumerator enumerator = this.m_A.FloatingShapesList.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				ShapeValue shapeValue = (ShapeValue)enumerator.Current;
				if (!shapeValue.IsInsideGroup || !RedactUtilities.IsPictureOrChart(shapeValue.Shape))
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
					if (ProgressBarLaunched)
					{
						base.Dispatcher.Invoke([SpecialName] () =>
						{
							this.m_A = RedactUtilities.ShowYesNoDialogue(System.Windows.Window.GetWindow(this), Conversions.ToString(this.m_C));
						});
					}
					else
					{
						this.m_A = RedactUtilities.ShowYesNoDialogue(null, Conversions.ToString(this.m_C));
					}
					return this.m_A;
				}
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_00c2;
				}
				continue;
				end_IL_00c2:
				break;
			}
		}
		return true;
	}

	private bool D()
	{
		if (this.m_A)
		{
			return true;
		}
		List<GroupedShapesValue> groupedFloatingShapesList = this.m_A.GroupedFloatingShapesList;
		groupedFloatingShapesList.AddRange(this.m_A.GroupedNonFloatingShapesList);
		using (List<GroupedShapesValue>.Enumerator enumerator = groupedFloatingShapesList.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				if (!enumerator.Current.HasPictureOrChart)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (ProgressBarLaunched)
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
						base.Dispatcher.Invoke([SpecialName] () =>
						{
							this.m_A = RedactUtilities.ShowYesNoDialogue(System.Windows.Window.GetWindow(this), Conversions.ToString(this.m_C));
						});
					}
					else
					{
						this.m_A = RedactUtilities.ShowYesNoDialogue(null, Conversions.ToString(this.m_C));
					}
					return this.m_A;
				}
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_00c2;
				}
				continue;
				end_IL_00c2:
				break;
			}
		}
		return true;
	}

	private bool B(DoWorkEventArgs A)
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
			if (this.m_A.CancellationPending)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						A.Cancel = true;
						return true;
					}
				}
			}
		}
		return false;
	}

	private void bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
	{
		pbRedact.Value = e.ProgressPercentage;
	}

	private void bgw_Completed(object sender, RunWorkerCompletedEventArgs e)
	{
		int num;
		if (!e.Cancelled)
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
			num = (Conversions.ToBoolean(this.m_B) ? 1 : 0);
		}
		else
		{
			num = 1;
		}
		object value = (byte)num != 0;
		base.Topmost = false;
		Hide();
		if (Conversions.ToBoolean(value))
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
			Forms.WarningMessage(System.Windows.Window.GetWindow(this), XC.A(1663));
		}
		this.m_A = null;
		Close();
	}

	private void H()
	{
		if (this.m_A == null)
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
			if (Operators.ConditionalCompareObjectEqual(this.m_A, 2, TextCompare: false))
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
				base.Dispatcher.Invoke([SpecialName] () =>
				{
					pbRedact.IsIndeterminate = false;
				}, DispatcherPriority.Background);
			}
			if (!Operators.ConditionalCompareObjectGreater(this.m_A, 1, TextCompare: false))
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
				int percentProgress = Conversions.ToInteger(Operators.MultiplyObject(Operators.DivideObject(Operators.SubtractObject(this.m_A, 1), this.m_A), 100));
				this.m_A.ReportProgress(percentProgress);
				base.Dispatcher.Invoke([SpecialName] () =>
				{
					txtRedact.Text = XC.A(18714) + this.m_A.ToString() + XC.A(13138) + this.m_A;
				}, DispatcherPriority.Background);
				return;
			}
		}
	}

	private void I()
	{
		this.m_A.ReportProgress(100);
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			txtRedact.Text = XC.A(15333);
		}, DispatcherPriority.Background);
	}

	private void wpfRedact_Closing(object sender, CancelEventArgs e)
	{
		if (this.m_A != null && this.m_A.IsBusy)
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
			this.m_A.CancelAsync();
			e.Cancel = true;
		}
		if (e.Cancel)
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
			this.m_A = null;
			this.m_A = null;
			ReleaseHelper.DoGarbageCollection();
			return;
		}
	}

	private void btnCancel_MouseEnter(object sender, MouseEventArgs e)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			btnCancel.Opacity = 1.0;
		}, DispatcherPriority.Background);
	}

	private void btnCancel_MouseLeave(object sender, MouseEventArgs e)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			btnCancel.Opacity = 0.6;
		}, DispatcherPriority.Background);
	}

	private void btnCancel_Click(object sender, RoutedEventArgs e)
	{
		try
		{
			if (!this.m_A.IsBusy)
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
				this.m_A.CancelAsync();
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

	private void Window_MouseDown(object sender, MouseButtonEventArgs e)
	{
		if (e.ChangedButton == MouseButton.Left)
		{
			DragMove();
		}
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_C)
		{
			this.m_C = true;
			Uri resourceLocator = new Uri(XC.A(18562), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
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
					btnCancel = (Button)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 2:
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				pbRedact = (ProgressBar)target;
				return;
			}
		case 3:
			txtRedact = (TextBlock)target;
			break;
		default:
			this.m_C = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}

	[SpecialName]
	[CompilerGenerated]
	private void J()
	{
		this.m_A = RedactUtilities.ShowYesNoDialogue(System.Windows.Window.GetWindow(this), Conversions.ToString(this.m_C));
	}

	[SpecialName]
	[CompilerGenerated]
	private void K()
	{
		this.m_A = RedactUtilities.ShowYesNoDialogue(System.Windows.Window.GetWindow(this), Conversions.ToString(this.m_C));
	}

	[SpecialName]
	[CompilerGenerated]
	private void L()
	{
		pbRedact.IsIndeterminate = false;
	}

	[SpecialName]
	[CompilerGenerated]
	private void M()
	{
		txtRedact.Text = XC.A(18714) + this.m_A.ToString() + XC.A(13138) + this.m_A;
	}

	[SpecialName]
	[CompilerGenerated]
	private void N()
	{
		txtRedact.Text = XC.A(15333);
	}

	[SpecialName]
	[CompilerGenerated]
	private void O()
	{
		btnCancel.Opacity = 1.0;
	}

	[SpecialName]
	[CompilerGenerated]
	private void P()
	{
		btnCancel.Opacity = 0.6;
	}
}
