using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Markup;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Macabacus_Word.Shapes;
using Macabacus_Word.TextOps.Redaction.Redactors;
using Macabacus_Word.TextOps.Redaction.Values;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.TextOps.Redaction.Process;

[DesignerGenerated]
public sealed class wpfFindAndRedact : System.Windows.Window, IComponentConnector
{
	[CompilerGenerated]
	internal sealed class SB
	{
		public string A;

		public wpfFindAndRedact A;

		public SB(SB A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			this.A.lblStatus.Text = XC.A(19247) + this.A + XC.A(44321);
		}
	}

	[CompilerGenerated]
	internal sealed class TB
	{
		public string A;

		public wpfFindAndRedact A;

		public TB(TB A)
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
		internal void A()
		{
			this.A.lblStatus.Text = XC.A(19247) + this.A + XC.A(44321);
		}
	}

	[CompilerGenerated]
	internal sealed class UB
	{
		public Exception A;

		public wpfFindAndRedact A;

		[SpecialName]
		internal void A()
		{
			Forms.ErrorMessage(System.Windows.Window.GetWindow(this.A), XC.A(44220) + this.A.Message);
		}
	}

	[CompilerGenerated]
	internal sealed class VB
	{
		public Exception A;

		public wpfFindAndRedact A;

		public VB(VB A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			Forms.ErrorMessage(System.Windows.Window.GetWindow(this.A), XC.A(44220) + this.A.Message);
		}
	}

	[CompilerGenerated]
	internal sealed class WB
	{
		public string A;

		public wpfFindAndRedact A;

		public WB(WB A)
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
		internal void A()
		{
			this.A.lblStatus.Text = XC.A(19247) + this.A + XC.A(44321);
		}
	}

	private BackgroundWorker m_A;

	private SelectionValue m_A;

	private string m_A;

	private int m_A;

	private bool m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("txtFind")]
	private System.Windows.Controls.TextBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lblSearching")]
	private TextBlock m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lblStatus")]
	private TextBlock m_B;

	[AccessedThroughProperty("btnRedact")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private System.Windows.Controls.Button m_B;

	private bool m_B;

	internal virtual System.Windows.Controls.TextBox txtFind
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

	internal virtual TextBlock lblSearching
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

	internal virtual TextBlock lblStatus
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

	internal virtual System.Windows.Controls.Button btnRedact
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
			RoutedEventHandler value2 = btnRedact_Click;
			System.Windows.Controls.Button button = this.m_A;
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

	internal virtual System.Windows.Controls.Button btnCancel
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
			RoutedEventHandler value2 = btnCancel_Click;
			System.Windows.Controls.Button button = this.m_B;
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
			this.m_B = value;
			button = this.m_B;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	public wpfFindAndRedact(SelectionValue selectionValue)
	{
		base.Closing += wpfRedact_Closing;
		this.m_A = null;
		this.m_A = "";
		this.m_A = 0;
		this.m_A = false;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		lblStatus.Text = "";
		lblSearching.Visibility = Visibility.Hidden;
		this.m_A = selectionValue;
	}

	private void wpfRedact_Closing(object sender, CancelEventArgs e)
	{
		try
		{
			this.m_A.CancelAsync();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void btnRedact_Click(object sender, RoutedEventArgs e)
	{
		bool flag = false;
		try
		{
			if (txtFind.Text.Length == 0)
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
						Forms.WarningMessage(XC.A(18913));
						return;
					}
				}
			}
			Document activeDocument = this.m_A.WdApp.ActiveDocument;
			if (!activeDocument.Saved)
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
				if (activeDocument.Path.Length > 0)
				{
					DialogResult dialogResult = Forms.YesNoCancelMessage(XC.A(18946));
					if (dialogResult != System.Windows.Forms.DialogResult.Cancel)
					{
						if (dialogResult == System.Windows.Forms.DialogResult.Yes)
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
							activeDocument.Save();
						}
					}
					else
					{
						flag = true;
					}
				}
			}
			activeDocument = null;
			if (!flag)
			{
				btnRedact.IsEnabled = false;
				btnCancel.Content = XC.A(19096);
				this.m_A = new BackgroundWorker();
				BackgroundWorker a = this.m_A;
				a.WorkerSupportsCancellation = true;
				a.WorkerReportsProgress = false;
				a.DoWork += bgw_DoWork;
				a.RunWorkerCompleted += bgw_RunWorkerCompleted;
				a.RunWorkerAsync();
				_ = null;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Exception A = ex2;
			base.Dispatcher.Invoke([SpecialName] () =>
			{
				Forms.ErrorMessage(System.Windows.Window.GetWindow(this), XC.A(44220) + A.Message);
			});
			clsReporting.LogException(A);
			ProjectData.ClearProjectError();
		}
		finally
		{
			Document activeDocument = null;
		}
	}

	private void bgw_DoWork(object sender, DoWorkEventArgs e)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			lblSearching.Visibility = Visibility.Visible;
			this.m_A = txtFind.Text.Trim();
		});
		Document activeDocument = this.m_A.WdApp.ActiveDocument;
		UndoRecord undo = RedactUtilities.BeginRedactionProcess(XC.A(980), this.m_A.WdApp);
		try
		{
			_ = activeDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.StoryType;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = activeDocument.StoryRanges.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					while (true)
					{
						if (this.m_A.CancellationPending)
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
							e.Cancel = true;
							break;
						}
						A(range, this.m_A, ref this.m_A, ref e);
						range = range.NextStoryRange;
						if (range != null)
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
						switch (3)
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
			VB a = default(VB);
			VB CS_0024_003C_003E8__locals5 = new VB(a);
			CS_0024_003C_003E8__locals5.A = this;
			Exception a2 = ex;
			CS_0024_003C_003E8__locals5.A = a2;
			base.Dispatcher.Invoke([SpecialName] () =>
			{
				Forms.ErrorMessage(System.Windows.Window.GetWindow(CS_0024_003C_003E8__locals5.A), XC.A(44220) + CS_0024_003C_003E8__locals5.A.Message);
			});
			clsReporting.LogException(CS_0024_003C_003E8__locals5.A);
			if (this.m_A != null)
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
				if (this.m_A.IsBusy)
				{
					e.Cancel = true;
					this.m_A.CancelAsync();
				}
			}
			this.m_A = true;
			base.Dispatcher.Invoke([SpecialName] () =>
			{
				Close();
			});
			ProjectData.ClearProjectError();
		}
		finally
		{
			RedactUtilities.CompleteRedactionProcess(XC.A(19105), isFindAndRedact: true, ref undo, this.m_A);
			undo = null;
			activeDocument = null;
		}
	}

	private void A(Range A, string B, ref int C, ref DoWorkEventArgs D)
	{
		WB a = default(WB);
		WB CS_0024_003C_003E8__locals4 = new WB(a);
		CS_0024_003C_003E8__locals4.A = this;
		if (this.B(A))
		{
			return;
		}
		clsUtilities.IsRangeInHeaderFooter(A);
		Find findObject = RedactUtilities.GetFindObject(A, B);
		_ = A.StoryType;
		checked
		{
			while (true)
			{
				Find find = findObject;
				object FindText = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchCase = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchWholeWord = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchWildcards = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchSoundsLike = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchAllWordForms = RuntimeHelpers.GetObjectValue(Missing.Value);
				object Forward = RuntimeHelpers.GetObjectValue(Missing.Value);
				object Wrap = RuntimeHelpers.GetObjectValue(Missing.Value);
				object Format = RuntimeHelpers.GetObjectValue(Missing.Value);
				object ReplaceWith = RuntimeHelpers.GetObjectValue(Missing.Value);
				object Replace = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchKashida = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchDiacritics = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchAlefHamza = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchControl = RuntimeHelpers.GetObjectValue(Missing.Value);
				if (!find.Execute(ref FindText, ref MatchCase, ref MatchWholeWord, ref MatchWildcards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl))
				{
					break;
				}
				if (this.m_A.CancellationPending)
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
					D.Cancel = true;
					break;
				}
				if (!clsUtilities.IsRangeInHeaderFooter(A))
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
					TextRedactor.RedactLineOfWordsInterop(A.Duplicate);
				}
				else
				{
					TextRedactor.RedactWordsInterop(A.Duplicate);
				}
				C++;
				CS_0024_003C_003E8__locals4.A = C.ToString();
				base.Dispatcher.Invoke([SpecialName] () =>
				{
					CS_0024_003C_003E8__locals4.A.lblStatus.Text = XC.A(19247) + CS_0024_003C_003E8__locals4.A + XC.A(44321);
				});
			}
			if (this.A(A))
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
				this.B(A, B, ref C, ref D);
				this.A(B, ref C, ref D);
			}
			findObject = null;
		}
	}

	private void A(string A, ref int B, ref DoWorkEventArgs C)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = this.m_A.WdApp.ActiveDocument.InlineShapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				InlineShape inlineShape = (InlineShape)enumerator.Current;
				if (inlineShape.Type != WdInlineShapeType.wdInlineShapeSmartArt)
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
				this.A(A, inlineShape.SmartArt, ref B, ref C);
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

	private void B(Range A, string B, ref int C, ref DoWorkEventArgs D)
	{
		if (!Helpers.A(A))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.ShapeRange.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Word.Shape shape = (Microsoft.Office.Interop.Word.Shape)enumerator.Current;
				if (shape.Type != MsoShapeType.msoGroup)
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
					this.A(shape, B, ref C, ref D);
					continue;
				}
				try
				{
					enumerator2 = shape.GroupItems.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Microsoft.Office.Interop.Word.Shape a = (Microsoft.Office.Interop.Word.Shape)enumerator2.Current;
						this.A(a, B, ref C, ref D);
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
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (3)
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

	private void B(string A, ref int B, ref DoWorkEventArgs C)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = this.m_A.WdApp.ActiveDocument.Shapes.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Word.Shape shape = (Microsoft.Office.Interop.Word.Shape)enumerator.Current;
				if (shape.Type != MsoShapeType.msoGroup)
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
					this.A(shape, A, ref B, ref C);
					continue;
				}
				try
				{
					enumerator2 = shape.GroupItems.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Microsoft.Office.Interop.Word.Shape a = (Microsoft.Office.Interop.Word.Shape)enumerator2.Current;
						this.A(a, A, ref B, ref C);
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
					switch (3)
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

	private void A(Microsoft.Office.Interop.Word.Shape A, string B, ref int C, ref DoWorkEventArgs D)
	{
		if (A.Type == MsoShapeType.msoSmartArt)
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
					this.A(B, A.SmartArt, ref C, ref D);
					return;
				}
			}
		}
		if (A.Type != MsoShapeType.msoAutoShape)
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
			if (A.Type != MsoShapeType.msoTextBox)
			{
				return;
			}
		}
		if (!clsUtilities.DoesAutoShapeHaveText(A))
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
			if (RedactUtilities.GetWordList(A.TextFrame.TextRange, trimList: true).Count() <= 0)
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
				this.B(B, A, ref C, ref D);
				return;
			}
		}
	}

	private void B(string A, Microsoft.Office.Interop.Word.Shape B, ref int C, ref DoWorkEventArgs D)
	{
		SB a = default(SB);
		SB CS_0024_003C_003E8__locals4 = new SB(a);
		CS_0024_003C_003E8__locals4.A = this;
		Range textRange = B.TextFrame.TextRange;
		Find findObject = RedactUtilities.GetFindObject(textRange, A);
		checked
		{
			while (true)
			{
				Find find = findObject;
				object FindText = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchCase = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchWholeWord = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchWildcards = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchSoundsLike = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchAllWordForms = RuntimeHelpers.GetObjectValue(Missing.Value);
				object Forward = RuntimeHelpers.GetObjectValue(Missing.Value);
				object Wrap = RuntimeHelpers.GetObjectValue(Missing.Value);
				object Format = RuntimeHelpers.GetObjectValue(Missing.Value);
				object ReplaceWith = RuntimeHelpers.GetObjectValue(Missing.Value);
				object Replace = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchKashida = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchDiacritics = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchAlefHamza = RuntimeHelpers.GetObjectValue(Missing.Value);
				object MatchControl = RuntimeHelpers.GetObjectValue(Missing.Value);
				if (find.Execute(ref FindText, ref MatchCase, ref MatchWholeWord, ref MatchWildcards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl))
				{
					if (this.m_A.CancellationPending)
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
						D.Cancel = true;
						break;
					}
					AutoShapeRedactor.RedactRangeAutoshape(B, textRange.Duplicate);
					C++;
					CS_0024_003C_003E8__locals4.A = C.ToString();
					base.Dispatcher.Invoke([SpecialName] () =>
					{
						CS_0024_003C_003E8__locals4.A.lblStatus.Text = XC.A(19247) + CS_0024_003C_003E8__locals4.A + XC.A(44321);
					});
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
				break;
			}
			findObject = null;
			textRange = null;
		}
	}

	private void A(string A, SmartArt B, ref int C, ref DoWorkEventArgs D)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = B.AllNodes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				SmartArtNode smartArtNode = (SmartArtNode)enumerator.Current;
				this.A((TextRange2)NewLateBinding.LateGet(NewLateBinding.LateGet(smartArtNode.Shapes.Cast<object>().ElementAtOrDefault(0), null, XC.A(19132), new object[0], null, null, null), null, XC.A(19153), new object[0], null, null, null), A, ref C, ref D);
			}
		}
		finally
		{
			if (enumerator is IDisposable)
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
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	private void A(TextRange2 A, string B, ref int C, ref DoWorkEventArgs D)
	{
		TB a = default(TB);
		TB CS_0024_003C_003E8__locals4 = new TB(a);
		CS_0024_003C_003E8__locals4.A = this;
		TextRange2 textRange = null;
		int num = 0;
		int length = B.Length;
		textRange = A.Find(B, 0, MsoTriState.msoFalse, MsoTriState.msoTrue);
		checked
		{
			while (textRange != null)
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
					if (textRange.Length == length)
					{
						if (this.m_A.CancellationPending)
						{
							D.Cancel = true;
							return;
						}
						num = textRange.Start + textRange.Length - 1;
						SmartArtRedactor.RedactRangeInSmartArt(textRange);
						C++;
						CS_0024_003C_003E8__locals4.A = C.ToString();
						base.Dispatcher.Invoke([SpecialName] () =>
						{
							CS_0024_003C_003E8__locals4.A.lblStatus.Text = XC.A(19247) + CS_0024_003C_003E8__locals4.A + XC.A(44321);
						});
						textRange = A.Find(B, num, MsoTriState.msoFalse, MsoTriState.msoTrue);
						break;
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
			}
		}
	}

	private bool A(Range A)
	{
		if (A.StoryType != WdStoryType.wdMainTextStory && A.StoryType != WdStoryType.wdEvenPagesHeaderStory)
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
			if (A.StoryType != WdStoryType.wdPrimaryHeaderStory)
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
				if (A.StoryType != WdStoryType.wdEvenPagesFooterStory)
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
					if (A.StoryType != WdStoryType.wdPrimaryFooterStory)
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
						if (A.StoryType != WdStoryType.wdFirstPageHeaderStory)
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
							if (A.StoryType != WdStoryType.wdFirstPageFooterStory)
							{
								return false;
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
					}
				}
			}
		}
		return true;
	}

	private bool B(Range A)
	{
		if (A.StoryType != WdStoryType.wdTextFrameStory)
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
			if (A.StoryType != WdStoryType.wdFootnoteSeparatorStory && A.StoryType != WdStoryType.wdFootnoteContinuationSeparatorStory)
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
				if (A.StoryType != WdStoryType.wdEndnoteSeparatorStory)
				{
					if (A.StoryType != WdStoryType.wdEndnoteContinuationSeparatorStory)
					{
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
			}
		}
		return true;
	}

	private void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
		if (!this.m_A && !e.Cancelled)
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
			base.DialogResult = true;
			Forms.InfoMessage(A());
		}
		if (this.m_A != null)
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
			this.m_A.Dispose();
			this.m_A = null;
		}
		this.m_A = null;
	}

	private void btnCancel_Click(object sender, RoutedEventArgs e)
	{
		if (this.m_A != null)
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
			if (this.m_A.IsBusy)
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
				this.m_A.CancelAsync();
			}
			if (this.m_A > 0)
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
				Forms.InfoMessage(A() + XC.A(19172));
			}
		}
		base.DialogResult = true;
	}

	private string A()
	{
		if (this.m_A > 0)
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
					return XC.A(19247) + this.m_A + XC.A(19266) + this.m_A + XC.A(19301);
				}
			}
		}
		return XC.A(19306);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_B)
		{
			this.m_B = true;
			Uri resourceLocator = new Uri(XC.A(19337), UriKind.Relative);
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
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					txtFind = (System.Windows.Controls.TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					lblSearching = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					lblStatus = (TextBlock)target;
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
					btnRedact = (System.Windows.Controls.Button)target;
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
					btnCancel = (System.Windows.Controls.Button)target;
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

	[SpecialName]
	[CompilerGenerated]
	private void A()
	{
		lblSearching.Visibility = Visibility.Visible;
		this.m_A = txtFind.Text.Trim();
	}

	[SpecialName]
	[CompilerGenerated]
	private void B()
	{
		Close();
	}
}
