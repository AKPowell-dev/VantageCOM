using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using A;
using ExcelAddIn1.Audit.Check.Analyses;
using ExcelAddIn1.Audit.Check.Helpers;
using ExcelAddIn1.Audit.Check.Observations;
using MacabacusMacros;
using MacabacusMacros.Xaml;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check;

public sealed class Analysis : RangeHelpers.IActionStatusUpdater
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<RB, bool> A;

		public static Func<ActionItem, bool> A;

		public static Func<ActionItem, bool> B;

		public static Func<ActionItem, bool> C;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(RB A)
		{
			ActionItem associatedAction = A.AssociatedAction;
			bool? obj;
			if (associatedAction == null)
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
				obj = null;
			}
			else
			{
				obj = associatedAction.ErroredOut;
			}
			return object.Equals(obj, true);
		}

		[SpecialName]
		internal bool A(ActionItem A)
		{
			return A.IsSkipped;
		}

		[SpecialName]
		internal bool B(ActionItem A)
		{
			return A.B();
		}

		[SpecialName]
		internal bool C(ActionItem A)
		{
			return A.IsCompletedFully;
		}
	}

	[CompilerGenerated]
	internal sealed class Z
	{
		public Dictionary<string, bool> A;

		public Z(Z A)
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
		internal void A(Microsoft.Office.Interop.Excel.Workbook A)
		{
			this.A[A.FullName] = A.Saved;
		}

		[SpecialName]
		internal void B(Microsoft.Office.Interop.Excel.Workbook A)
		{
			if (!this.A.TryGetValue(A.FullName, out var value))
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
				if (A.Saved == value)
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
					A.Saved = value;
					return;
				}
			}
		}
	}

	[CompilerGenerated]
	internal sealed class AB
	{
		public StringBuilder A;

		[SpecialName]
		internal void A(int A, string B)
		{
			if (A < 1)
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
				if (this.A.Length > 0)
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
					this.A.AppendLine();
				}
				this.A.Append(string.Format(VH.A(211178), A, B));
				return;
			}
		}
	}

	[CompilerGenerated]
	internal sealed class BB
	{
		public Action A;

		public BB(BB A)
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
		internal void A()
		{
			this.A();
		}
	}

	[CompilerGenerated]
	private List<Observation> m_A;

	[CompilerGenerated]
	private List<int> m_A;

	[CompilerGenerated]
	private string m_A;

	[CompilerGenerated]
	private string m_B;

	[CompilerGenerated]
	private string m_C;

	[CompilerGenerated]
	private bool m_A;

	[CompilerGenerated]
	private GB m_A;

	[CompilerGenerated]
	private Dictionary<string, Dictionary<string, List<Range>>> m_A;

	[CompilerGenerated]
	private FB m_A;

	private readonly BackgroundWorker m_A;

	private readonly StatusKeeper m_A;

	private readonly Action<bool> m_A;

	private long? m_A;

	private bool m_B;

	private int m_A;

	private int m_B;

	internal List<Observation> Observations
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

	internal List<int> PaletteColors
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

	internal string CancelSkipText
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

	internal string CheckExceptionsText
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

	internal string CheckExceptionsTextDetailed
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	internal bool HasGeneralException
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

	internal GB PrecRetriever
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

	internal Dictionary<string, Dictionary<string, List<Range>>> DictWsFormulas
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

	internal FB ParenthesisPairs
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

	internal int A => this.m_A.ActionLevel;

	internal int B
	{
		get
		{
			return this.m_A;
		}
		set
		{
			if (object.Equals(this.m_A, value))
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
				this.m_A = value;
				this.m_A.NumChecksToRun = value;
				return;
			}
		}
	}

	internal int C
	{
		get
		{
			return this.m_B;
		}
		set
		{
			if (object.Equals(this.m_B, value))
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
				this.m_B = value;
				this.m_A.CurrentCheckNum = value;
				return;
			}
		}
	}

	internal Analysis(Settings A, bool B, ref BackgroundWorker C, StatusKeeper D, Action<bool> E)
	{
		Z a = default(Z);
		Z CS_0024_003C_003E8__locals3 = new Z(a);
		base._002Ector();
		PaletteColors = null;
		PrecRetriever = new GB();
		DictWsFormulas = new Dictionary<string, Dictionary<string, List<Range>>>();
		ParenthesisPairs = new FB(this);
		this.m_A = C;
		this.m_A = D;
		this.m_A.OnStatusChange = [SpecialName] () =>
		{
			this.C();
		};
		this.m_A = E;
		PaletteColors = new List<int>();
		try
		{
			foreach (PaletteColor item in clsColors.ColorPalette)
			{
				PaletteColors.Add(ColorTranslator.ToOle(clsColors.RGB2Color(item.RGB)));
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		Observations = new List<Observation>();
		Application application = MH.A.Application;
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application.ActiveWorkbook;
		Microsoft.Office.Interop.Excel.Sheets selectedSheets = application.ActiveWindow.SelectedSheets;
		object objectValue = RuntimeHelpers.GetObjectValue(application.ActiveSheet);
		Range range = null;
		Range activeCell = application.ActiveCell;
		if (application.Selection is Range)
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
			range = (Range)application.Selection;
		}
		application.ScreenUpdating = false;
		List<RB> list = new List<RB>();
		this.m_A.F();
		CS_0024_003C_003E8__locals3.A = new Dictionary<string, bool>();
		Microsoft.Office.Interop.Excel.Sheets e;
		try
		{
			OB.A(application, [SpecialName] (Microsoft.Office.Interop.Excel.Workbook workbook) =>
			{
				CS_0024_003C_003E8__locals3.A[workbook.FullName] = workbook.Saved;
			});
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
				list.AddRange(ExcelApplication.A(this, A, application));
				list.AddRange(ExcelAddIn1.Audit.Check.Analyses.Workbook.A(this, A, activeWorkbook));
				e = activeWorkbook.Sheets;
			}
			else
			{
				e = selectedSheets;
			}
			list.AddRange(ExcelAddIn1.Audit.Check.Analyses.Worksheet.A(this, A, B, PaletteColors));
			list.AddRange(ChartSheet.A(this, A));
			RB.B(this, list, application, activeWorkbook, e);
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			HasGeneralException = true;
			this.m_A.B(ex4.Message);
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
		finally
		{
			OB.A(application, [SpecialName] (Microsoft.Office.Interop.Excel.Workbook workbook) =>
			{
				if (CS_0024_003C_003E8__locals3.A.TryGetValue(workbook.FullName, out var value))
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
							if (workbook.Saved != value)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										workbook.Saved = value;
										return;
									}
								}
							}
							return;
						}
					}
				}
			});
			this.m_A.G();
		}
		this.A(list);
		this.A();
		if (selectedSheets.Count > 1)
		{
			selectedSheets.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		if (objectValue is Microsoft.Office.Interop.Excel.Worksheet)
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
			((Microsoft.Office.Interop.Excel.Worksheet)objectValue).Activate();
		}
		else if (objectValue is Chart)
		{
			((Chart)objectValue).Activate();
		}
		if (range != null)
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
			range.Select();
		}
		activeCell.Activate();
		application.ScreenUpdating = true;
		e = null;
		selectedSheets = null;
		objectValue = null;
		range = null;
		activeCell = null;
		activeWorkbook = null;
		application = null;
	}

	private void A(List<RB> A)
	{
		Func<RB, bool> predicate;
		if (_Closure_0024__.A == null)
		{
			predicate = (_Closure_0024__.A = [SpecialName] (RB rB) =>
			{
				ActionItem associatedAction = rB.AssociatedAction;
				bool? obj;
				if (associatedAction == null)
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
					obj = null;
				}
				else
				{
					obj = associatedAction.ErroredOut;
				}
				return object.Equals(obj, true);
			});
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
		List<RB> list = A.Where(predicate).ToList();
		if (list.Count == 0)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					CheckExceptionsText = "";
					CheckExceptionsTextDetailed = "";
					return;
				}
			}
		}
		CheckExceptionsText = string.Format(VH.A(6914), list.Count);
		StringBuilder stringBuilder = new StringBuilder();
		using (List<RB>.Enumerator enumerator = list.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				RB current = enumerator.Current;
				StringBuilder stringBuilder2 = stringBuilder.AppendLine().AppendLine().Append(VH.A(7100))
					.Append(current.CheckDesc)
					.Append(VH.A(7115))
					.AppendLine()
					.Append(VH.A(7120));
				Exception exception = current.AssociatedAction.Exception;
				object value;
				if (exception == null)
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
					value = null;
				}
				else
				{
					value = exception.GetType().Name;
				}
				stringBuilder2.Append((string)value).Append(VH.A(7123)).Append(current.AssociatedAction.Exception?.Message);
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_0181;
				}
				continue;
				end_IL_0181:
				break;
			}
		}
		CheckExceptionsTextDetailed = stringBuilder.ToString().TrimStart();
	}

	private void A()
	{
		StatusKeeper a = this.m_A;
		int num;
		if (a == null)
		{
			num = 0;
		}
		else
		{
			ObservableCollection<ActionItem> allItems = a.AllItems;
			Func<ActionItem, bool> predicate;
			if (_Closure_0024__.A == null)
			{
				predicate = (_Closure_0024__.A = [SpecialName] (ActionItem actionItem) => actionItem.IsSkipped);
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
			num = allItems.Where(predicate).Count();
		}
		int num2 = num;
		StatusKeeper a2 = this.m_A;
		int num3;
		if (a2 == null)
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
			num3 = 0;
		}
		else
		{
			num3 = a2.AllItems.Where([SpecialName] (ActionItem actionItem) => actionItem.B()).Count();
		}
		int num4 = num3;
		if (num2 == 0 && num4 == 0)
		{
			CancelSkipText = "";
			return;
		}
		StatusKeeper a3 = this.m_A;
		int num5;
		if (a3 == null)
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
			num5 = 0;
		}
		else
		{
			ObservableCollection<ActionItem> allItems2 = a3.AllItems;
			Func<ActionItem, bool> predicate2;
			if (_Closure_0024__.C == null)
			{
				predicate2 = (_Closure_0024__.C = [SpecialName] (ActionItem actionItem) => actionItem.IsCompletedFully);
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
				predicate2 = _Closure_0024__.C;
			}
			num5 = allItems2.Where(predicate2).Count();
		}
		int cnt = num5;
		StringBuilder A = new StringBuilder();
		if (this.A(A: false))
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
			A.Append(VH.A(7128));
		}
		B<int, string> obj = [SpecialName] (int num6, string B) =>
		{
			if (num6 >= 1)
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
						if (A.Length > 0)
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
							A.AppendLine();
						}
						A.Append(string.Format(VH.A(211178), num6, B));
						return;
					}
				}
			}
		};
		obj(cnt, VH.A(7213));
		obj(num2, VH.A(7232));
		obj(num4, VH.A(7259));
		CancelSkipText = A.ToString();
	}

	internal void B()
	{
		this.m_A.SkipIsOn = false;
	}

	internal bool A()
	{
		C();
		if (!this.m_A.CancellationPending)
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
					return this.m_A.SkipIsOn;
				}
			}
		}
		return true;
	}

	internal bool A(bool A)
	{
		if (A)
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
			C();
		}
		return this.m_A.CancellationPending;
	}

	private void C()
	{
		long num = this.m_A.A();
		if (this.m_A.HasValue)
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
			if (checked(num - this.m_A.Value) < 300)
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
				break;
			}
		}
		this.m_A = num;
		XamlUtilities.XamlDoEvents();
	}

	internal void B(List<RB> A)
	{
		foreach (RB item in A)
		{
			this.m_A.A(item);
		}
	}

	internal int A(RB A, long B = 1L)
	{
		int a = this.A;
		this.m_A.A(A.CheckDesc, B);
		this.m_A.B(A);
		return a;
	}

	internal int A(string A, long B = 1L)
	{
		int a = this.A;
		ActionStarted(A, B);
		return a;
	}

	internal void ActionStarted(string actionDesc, long numItems = 1L)
	{
		this.m_A.A(actionDesc, numItems);
	}

	void RangeHelpers.IActionStatusUpdater.ActionStarted(string actionDesc, long numItems = 1L)
	{
		//ILSpy generated this explicit interface implementation from .override directive in ActionStarted
		this.ActionStarted(actionDesc, numItems);
	}

	internal ActionItem A()
	{
		return this.m_A.A();
	}

	internal void A(string A = "")
	{
		this.m_A.A(A);
	}

	private void D()
	{
		this.m_A.D();
	}

	internal bool ItemCancelled(string itemDesc = "")
	{
		if (this.A())
		{
			return true;
		}
		A(itemDesc);
		return false;
	}

	bool RangeHelpers.IActionStatusUpdater.ItemCancelled(string itemDesc = "")
	{
		//ILSpy generated this explicit interface implementation from .override directive in ItemCancelled
		return this.ItemCancelled(itemDesc);
	}

	internal void ActionEnded()
	{
		this.m_A.A(this.A());
	}

	void RangeHelpers.IActionStatusUpdater.ActionEnded()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ActionEnded
		this.ActionEnded();
	}

	internal void A(string A, Action B)
	{
		BB a = default(BB);
		BB CS_0024_003C_003E8__locals2 = new BB(a);
		CS_0024_003C_003E8__locals2.A = B;
		int a2 = this.A;
		bool a3 = this.B(A: true);
		try
		{
			ActionStarted(A, 1L);
			CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
			Task task = new Task([SpecialName] () =>
			{
				CS_0024_003C_003E8__locals2.A();
			}, cancellationTokenSource.Token, TaskCreationOptions.LongRunning);
			task.Start();
			while (!task.IsCompleted)
			{
				D();
				Thread.Sleep(200);
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
				ActionEnded();
				return;
			}
		}
		finally
		{
			this.A(a2);
			this.B(a3);
		}
	}

	private bool B(bool A)
	{
		bool b = this.m_B;
		this.m_B = A;
		this.m_A?.Invoke(A);
		return b;
	}

	internal void A(int A)
	{
		this.m_A.A(A, this.A());
	}

	[SpecialName]
	[CompilerGenerated]
	private void E()
	{
		C();
	}
}
