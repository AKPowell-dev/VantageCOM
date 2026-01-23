using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
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
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.DocBuilder;

[DesignerGenerated]
public sealed class wpfDocWizard : System.Windows.Controls.UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Choice, bool> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(Choice A)
		{
			return A.IsChecked;
		}
	}

	[CompilerGenerated]
	internal sealed class YB
	{
		public string A;

		public string B;

		public Func<BaseQuestion, bool> A;

		public YB(YB A)
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
				B = A.B;
				return;
			}
		}

		[SpecialName]
		internal bool A(BaseQuestion A)
		{
			if (Operators.CompareString(A.ContentControl.Tag, this.A, TextCompare: false) == 0)
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
						return Operators.CompareString(A.ContentControl.ID, B, TextCompare: false) != 0;
					}
				}
			}
			return false;
		}
	}

	[CompilerGenerated]
	internal sealed class ZB
	{
		public string A;

		public string B;

		public ZB(ZB A)
		{
			if (A != null)
			{
				this.A = A.A;
				B = A.B;
			}
		}

		[SpecialName]
		internal bool A(BaseQuestion A)
		{
			if (Operators.CompareString(A.ContentControl.Tag, this.A, TextCompare: false) == 0)
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
						return Operators.CompareString(A.ContentControl.ID, B, TextCompare: false) != 0;
					}
				}
			}
			return false;
		}
	}

	[CompilerGenerated]
	internal sealed class AC
	{
		public string A;

		public Func<BaseQuestion, bool> A;

		public AC(AC A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal bool A(BaseQuestion A)
		{
			return Operators.CompareString(A.ContentControl.ID, this.A, TextCompare: false) == 0;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private ObservableCollection<BaseQuestion> m_A;

	private bool m_A;

	public ObservableCollection<BaseQuestion> Questions
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(XC.A(21881));
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
	}

	public wpfDocWizard()
	{
		base.Unloaded += wpfDocWizard_Unloaded;
		this.m_A = null;
		InitializeComponent();
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
			switch (5)
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

	public void LoadQuestionnaire(List<BaseQuestion> listQuestions)
	{
		Questions = new ObservableCollection<BaseQuestion>(listQuestions);
	}

	private void ReloadQuestionnaire(object sender, RoutedEventArgs e)
	{
		try
		{
			Document activeDocument = PC.A.Application.ActiveDocument;
			Dictionary<Microsoft.Office.Interop.Word.ContentControl, string> listAutoPopFields = null;
			LoadQuestionnaire(Base.GetQuestions(activeDocument, ref listAutoPopFields));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
	}

	private void wpfDocWizard_Unloaded(object sender, RoutedEventArgs e)
	{
		Questions = null;
	}

	private void OptionSelected(object sender, RoutedEventArgs e)
	{
		Choice choice = (Choice)((System.Windows.Controls.RadioButton)sender).DataContext;
		BaseQuestion question = choice.Question;
		question.ApplyButtonVisibility = Visibility.Visible;
		if (choice.Question is YesNoQuestion)
		{
			C(question.ContentControl);
		}
		else
		{
			C(choice.ContentControl);
		}
		question = null;
		choice = null;
	}

	private void TextInputted(object sender, TextChangedEventArgs e)
	{
		YB a = default(YB);
		YB CS_0024_003C_003E8__locals7 = new YB(a);
		System.Windows.Controls.TextBox obj = (System.Windows.Controls.TextBox)sender;
		BaseQuestion baseQuestion = (BaseQuestion)obj.DataContext;
		string text = obj.Text;
		if (text.Length > 0)
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
			baseQuestion.ApplyButtonVisibility = Visibility.Visible;
		}
		else
		{
			baseQuestion.ApplyButtonVisibility = Visibility.Hidden;
		}
		if (obj.IsKeyboardFocused)
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
			CS_0024_003C_003E8__locals7.A = baseQuestion.ContentControl.Tag;
			CS_0024_003C_003E8__locals7.B = baseQuestion.ContentControl.ID;
			try
			{
				IEnumerator<BaseQuestion> enumerator = default(IEnumerator<BaseQuestion>);
				try
				{
					ObservableCollection<BaseQuestion> questions = Questions;
					Func<BaseQuestion, bool> predicate;
					if (CS_0024_003C_003E8__locals7.A != null)
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
						predicate = CS_0024_003C_003E8__locals7.A;
					}
					else
					{
						predicate = (CS_0024_003C_003E8__locals7.A = [SpecialName] (BaseQuestion A) =>
						{
							if (Operators.CompareString(A.ContentControl.Tag, CS_0024_003C_003E8__locals7.A, TextCompare: false) == 0)
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
										return Operators.CompareString(A.ContentControl.ID, CS_0024_003C_003E8__locals7.B, TextCompare: false) != 0;
									}
								}
							}
							return false;
						});
					}
					enumerator = questions.Where(predicate).GetEnumerator();
					while (enumerator.MoveNext())
					{
						baseQuestion = enumerator.Current;
						if (baseQuestion is TextInput)
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
							((TextInput)baseQuestion).Text = text;
						}
						else
						{
							((AutofillField)baseQuestion).Text = text;
						}
					}
				}
				finally
				{
					enumerator?.Dispose();
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		baseQuestion = null;
	}

	private void ApplySingleResponse(object sender, RoutedEventArgs e)
	{
		if (Licensing.AllowDocBuilderOperation())
		{
			Microsoft.Office.Interop.Word.Application application = PC.A.Application;
			UndoRecord undoRecord = application.UndoRecord;
			BaseQuestion baseQuestion = (BaseQuestion)((System.Windows.Controls.Button)sender).DataContext;
			C(baseQuestion.ContentControl);
			undoRecord.StartCustomRecord(XC.A(21825));
			application.ScreenUpdating = false;
			A(baseQuestion);
			application.ScreenUpdating = true;
			undoRecord.EndCustomRecord();
			A(baseQuestion, B: true);
			undoRecord = null;
			baseQuestion = null;
		}
	}

	private void ApplyAllResponses(object sender, RoutedEventArgs e)
	{
		if (!Licensing.AllowDocBuilderOperation())
		{
			return;
		}
		checked
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
				Microsoft.Office.Interop.Word.Application application = PC.A.Application;
				UndoRecord undoRecord = application.UndoRecord;
				int num = 0;
				undoRecord.StartCustomRecord(XC.A(21825));
				application.ScreenUpdating = false;
				for (int i = Questions.Count - 1; i >= 0; i += -1)
				{
					try
					{
						if (Questions[i].ApplyButtonVisibility != Visibility.Visible)
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
							A(Questions[i]);
							A(Questions[i], B: false);
							num++;
							break;
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				application.ScreenUpdating = true;
				undoRecord.EndCustomRecord();
				application = null;
				undoRecord = null;
				if (num > 0)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							Forms.SuccessMessage(XC.A(21900) + num + XC.A(21917));
							return;
						}
					}
				}
				Forms.WarningMessage(XC.A(21940));
				return;
			}
		}
	}

	private void A(BaseQuestion A)
	{
		try
		{
			if (A is YesNoQuestion)
			{
				if (((YesNoQuestion)A).Choices[0].IsChecked)
				{
					this.A((YesNoQuestion)A, B: false);
				}
				else
				{
					this.A((YesNoQuestion)A, B: true);
				}
				return;
			}
			if (A is MultipleChoice)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
					{
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						MultipleChoice a = (MultipleChoice)A;
						List<Choice> choices = ((MultipleChoice)A).Choices;
						Func<Choice, bool> predicate;
						if (_Closure_0024__.A == null)
						{
							predicate = (_Closure_0024__.A = [SpecialName] (Choice choice) => choice.IsChecked);
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
							predicate = _Closure_0024__.A;
						}
						this.A(a, choices.First(predicate).ContentControl.ID);
						return;
					}
					}
				}
			}
			if (A is TextInput)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						this.A((TextInput)A, ((TextInput)A).Text);
						return;
					}
				}
			}
			this.A((AutofillField)A, ((AutofillField)A).Text);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	private void A(YesNoQuestion A, bool B)
	{
		this.A(A.ContentControl);
		Microsoft.Office.Interop.Word.ContentControl contentControl = A.ContentControl;
		if (B)
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
			Range range = contentControl.Range;
			this.B(A.ContentControl);
			contentControl.Delete(DeleteContents: true);
			this.A(range);
			range = null;
		}
		else
		{
			contentControl.Delete();
		}
		contentControl = null;
	}

	private void A(MultipleChoice A, string B)
	{
		Range range = A.ContentControl.Range;
		bool flag = false;
		using (List<Choice>.Enumerator enumerator = A.Choices.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				Choice current = enumerator.Current;
				this.A(current.ContentControl);
				Microsoft.Office.Interop.Word.ContentControl contentControl = current.ContentControl;
				if (Operators.CompareString(contentControl.ID, B, TextCompare: false) != 0)
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
					Range range2 = contentControl.Range;
					this.B(current.ContentControl);
					Range range3 = range2;
					object Direction = WdCollapseDirection.wdCollapseEnd;
					range3.Collapse(ref Direction);
					Range range4 = range2;
					Direction = WdUnits.wdCharacter;
					object Count = 1;
					if (Operators.CompareString(range4.Next(ref Direction, ref Count).Text, string.Empty, TextCompare: false) == 0)
					{
						flag = true;
					}
					contentControl.Delete(DeleteContents: true);
					this.A(range2);
					range2 = null;
				}
				else
				{
					contentControl.Delete();
				}
				contentControl = null;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_00f5;
				}
				continue;
				end_IL_00f5:
				break;
			}
		}
		this.A(A.ContentControl);
		A.ContentControl.Delete();
		Range last;
		if (flag)
		{
			Range range5 = range;
			object Count = WdCollapseDirection.wdCollapseEnd;
			range5.Collapse(ref Count);
			last = range.Characters.Last;
			string text = last.Text;
			if (Operators.CompareString(text, XC.A(17685), TextCompare: false) != 0)
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
				if (Operators.CompareString(text, XC.A(17685), TextCompare: false) != 0 && Operators.CompareString(text, XC.A(18455), TextCompare: false) != 0)
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
					if (Operators.CompareString(text, XC.A(21985), TextCompare: false) != 0)
					{
						goto IL_01f9;
					}
				}
			}
			Range range6 = last;
			Count = RuntimeHelpers.GetObjectValue(Missing.Value);
			object Direction = RuntimeHelpers.GetObjectValue(Missing.Value);
			range6.Delete(ref Count, ref Direction);
			goto IL_01f9;
		}
		goto IL_01fc;
		IL_01f9:
		last = null;
		goto IL_01fc;
		IL_01fc:
		range = null;
	}

	private void A(TextInput A, string B)
	{
		ZB a = default(ZB);
		ZB CS_0024_003C_003E8__locals4 = new ZB(a);
		CS_0024_003C_003E8__locals4.A = A.ContentControl.Tag;
		CS_0024_003C_003E8__locals4.B = A.ContentControl.ID;
		List<BaseQuestion> list = null;
		this.B(A.ContentControl);
		A.Apply(B);
		try
		{
			list = Questions.Where([SpecialName] (BaseQuestion baseQuestion) =>
			{
				if (Operators.CompareString(baseQuestion.ContentControl.Tag, CS_0024_003C_003E8__locals4.A, TextCompare: false) == 0)
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
							return Operators.CompareString(baseQuestion.ContentControl.ID, CS_0024_003C_003E8__locals4.B, TextCompare: false) != 0;
						}
					}
				}
				return false;
			}).ToList();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (list != null)
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
			if (list.Count > 0)
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
				if (System.Windows.Forms.MessageBox.Show(XC.A(21988) + list.Count + XC.A(22065), XC.A(2438), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
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
					using List<BaseQuestion>.Enumerator enumerator = list.GetEnumerator();
					while (enumerator.MoveNext())
					{
						BaseQuestion current = enumerator.Current;
						this.B(current.ContentControl);
						((TextInput)current).Apply(B);
						this.A(current, B: false);
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_013b;
						}
						continue;
						end_IL_013b:
						break;
					}
				}
			}
		}
		list = null;
	}

	private void A(AutofillField A, string B)
	{
		this.B(A.ContentControl);
		A.Apply(B);
	}

	private void A(Microsoft.Office.Interop.Word.ContentControl A)
	{
		A.LockContents = false;
		A.LockContentControl = false;
	}

	private void A(Range A)
	{
		object Unit = WdUnits.wdCharacter;
		A.Expand(ref Unit);
		string text = A.Text;
		if (Operators.CompareString(text, XC.A(17685), TextCompare: false) != 0 && Operators.CompareString(text, XC.A(17685), TextCompare: false) != 0)
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
			if (Operators.CompareString(text, XC.A(18455), TextCompare: false) != 0)
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
				if (Operators.CompareString(text, XC.A(21985), TextCompare: false) != 0)
				{
					return;
				}
			}
		}
		Unit = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Count = RuntimeHelpers.GetObjectValue(Missing.Value);
		A.Delete(ref Unit, ref Count);
	}

	private void B(Microsoft.Office.Interop.Word.ContentControl A)
	{
		AC a = default(AC);
		AC CS_0024_003C_003E8__locals5 = new AC(a);
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Range.ContentControls.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Word.ContentControl contentControl = (Microsoft.Office.Interop.Word.ContentControl)enumerator.Current;
				CS_0024_003C_003E8__locals5.A = contentControl.ID;
				try
				{
					ObservableCollection<BaseQuestion> questions = Questions;
					Func<BaseQuestion, bool> predicate;
					if (CS_0024_003C_003E8__locals5.A != null)
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
						predicate = CS_0024_003C_003E8__locals5.A;
					}
					else
					{
						predicate = (CS_0024_003C_003E8__locals5.A = [SpecialName] (BaseQuestion baseQuestion) => Operators.CompareString(baseQuestion.ContentControl.ID, CS_0024_003C_003E8__locals5.A, TextCompare: false) == 0);
					}
					this.A(questions.First(predicate), B: false);
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

	private void A(BaseQuestion A, bool B)
	{
		if (B)
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
					Questions.Remove(A);
					return;
				}
			}
		}
		Questions.Remove(A);
	}

	private void NavigateQuestion(object sender, MouseButtonEventArgs e)
	{
		BaseQuestion baseQuestion = (BaseQuestion)((TextBlock)sender).DataContext;
		C(baseQuestion.ContentControl);
		baseQuestion = null;
	}

	private void NavigateTextBox(object sender, KeyboardFocusChangedEventArgs e)
	{
		BaseQuestion baseQuestion = (BaseQuestion)((System.Windows.Controls.TextBox)sender).DataContext;
		C(baseQuestion.ContentControl);
		baseQuestion = null;
	}

	private void C(Microsoft.Office.Interop.Word.ContentControl A)
	{
		Microsoft.Office.Interop.Word.Application application = PC.A.Application;
		bool flag = false;
		try
		{
			Microsoft.Office.Interop.Word.Window activeWindow = application.ActiveWindow;
			Range range = A.Range;
			object Start = RuntimeHelpers.GetObjectValue(Missing.Value);
			activeWindow.ScrollIntoView(range, ref Start);
			new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(1839)).RemoveEventHandler(application, new ApplicationEvents4_WindowSelectionChangeEventHandler(clsRibbon.Application_WindowSelectionChange));
			flag = true;
			A.Range.Select();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		finally
		{
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(1839)).AddEventHandler(application, new ApplicationEvents4_WindowSelectionChangeEventHandler(clsRibbon.Application_WindowSelectionChange));
			}
			application = null;
		}
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (!this.m_A)
		{
			this.m_A = true;
			Uri resourceLocator = new Uri(XC.A(22152), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		switch (connectionId)
		{
		case 7:
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
				((System.Windows.Controls.Button)target).Click += ApplyAllResponses;
				return;
			}
		case 8:
			((System.Windows.Controls.Button)target).Click += ReloadQuestionnaire;
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

	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
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
			eventSetter.Event = UIElement.MouseLeftButtonUpEvent;
			eventSetter.Handler = new MouseButtonEventHandler(NavigateQuestion);
			((Style)target).Setters.Add(eventSetter);
		}
		if (connectionId == 2)
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
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = System.Windows.Controls.Primitives.ButtonBase.ClickEvent;
			eventSetter.Handler = new RoutedEventHandler(ApplySingleResponse);
			((Style)target).Setters.Add(eventSetter);
		}
		if (connectionId == 3)
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
			((System.Windows.Controls.RadioButton)target).Checked += OptionSelected;
		}
		if (connectionId == 4)
		{
			((System.Windows.Controls.RadioButton)target).Checked += OptionSelected;
		}
		if (connectionId == 5)
		{
			((System.Windows.Controls.TextBox)target).TextChanged += TextInputted;
			((System.Windows.Controls.TextBox)target).GotKeyboardFocus += NavigateTextBox;
		}
		if (connectionId != 6)
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
			((System.Windows.Controls.TextBox)target).TextChanged += TextInputted;
			((System.Windows.Controls.TextBox)target).GotKeyboardFocus += NavigateTextBox;
			return;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
