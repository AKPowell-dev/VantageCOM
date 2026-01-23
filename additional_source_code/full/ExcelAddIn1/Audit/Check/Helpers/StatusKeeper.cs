using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using A;
using MacabacusMacros.UI;
using MacabacusMacros.UI.FormsExtensions;
using MacabacusMacros.Xaml.Geometry;

namespace ExcelAddIn1.Audit.Check.Helpers;

public sealed class StatusKeeper : INotifyPropertyChanged
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<ActionItem, bool> A;

		public static Func<ActionItem, bool> B;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(ActionItem A)
		{
			return A.Status == CB.E;
		}

		[SpecialName]
		internal bool B(ActionItem A)
		{
			return new CB[3]
			{
				CB.B,
				CB.F,
				CB.A
			}.Contains(A.Status);
		}
	}

	[CompilerGenerated]
	internal sealed class MB
	{
		public string A;

		[SpecialName]
		internal bool A()
		{
			return UIFormsExtensions.AskYesNo((Window)null, this.A, true, true);
		}
	}

	[CompilerGenerated]
	internal sealed class NB
	{
		public string A;

		[SpecialName]
		internal bool A()
		{
			Forms.ErrorMessage(this.A);
			return true;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private readonly List<ActionItem> m_A;

	private readonly List<ActionItem> m_B;

	private readonly Stopwatch m_A;

	[CompilerGenerated]
	private ObservableCollection<ActionItem> m_A;

	private readonly Action<List<ActionItem>> m_A;

	[CompilerGenerated]
	private Action m_A;

	private int m_A;

	private Arc m_A;

	[CompilerGenerated]
	private double m_A;

	private bool m_A;

	private int m_B;

	private int m_C;

	private string m_A;

	private string m_B;

	private bool m_B;

	public ObservableCollection<ActionItem> AllItems
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
	}

	internal Action OnStatusChange
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

	private int SkippedCount
	{
		get
		{
			return this.m_A;
		}
		set
		{
			A(ref this.m_A, value, C: false, VH.A(8431));
		}
	}

	public Arc ProgressArc
	{
		get
		{
			return this.m_A;
		}
		set
		{
			A(ref this.m_A, value, C: false, VH.A(7274));
		}
	}

	internal double MinimumPercToShow
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

	public bool ShouldHidePerc
	{
		get
		{
			return this.m_A;
		}
		set
		{
			A(ref this.m_A, value, C: false, VH.A(7403));
		}
	}

	internal int NumChecksToRun
	{
		get
		{
			return this.m_B;
		}
		set
		{
			if (A(ref this.m_B, value, C: false, VH.A(8456)))
			{
				C();
				B();
			}
		}
	}

	internal int CurrentCheckNum
	{
		get
		{
			return this.m_C;
		}
		set
		{
			if (A(ref this.m_C, value, C: false, VH.A(8485)))
			{
				C();
				B();
			}
		}
	}

	public string HeaderInfo
	{
		get
		{
			return this.m_A;
		}
		set
		{
			A(ref this.m_A, value, C: false, VH.A(8516));
		}
	}

	internal int ActionLevel => this.m_A.Count;

	public string TotalTimeStr
	{
		get
		{
			return this.m_B;
		}
		set
		{
			A(ref this.m_B, value, C: false, VH.A(8593));
		}
	}

	internal bool SkipIsOn
	{
		get
		{
			return this.m_B;
		}
		set
		{
			if (!A(ref this.m_B, value, C: false, VH.A(8651)))
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
				JH.A();
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
		}
	}

	internal StatusKeeper(Action<List<ActionItem>> A)
	{
		//IL_0065: Unknown result type (might be due to invalid IL or missing references)
		//IL_006f: Expected O, but got Unknown
		this.m_A = new List<ActionItem>();
		this.m_B = new List<ActionItem>();
		this.m_A = new Stopwatch();
		this.m_A = new ObservableCollection<ActionItem>();
		MinimumPercToShow = 0.1;
		this.m_A = A;
		ProgressArc = new Arc(2.0, 2.0, 24.0, true);
	}

	public void NotifyPropertyChanged(string propertyName)
	{
		PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
		if (propertyChangedEventHandler == null)
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
			propertyChangedEventHandler(this, new PropertyChangedEventArgs(propertyName));
			return;
		}
	}

	private bool A<A>(ref A A, A B, bool C = false, [CallerMemberName] string D = null)
	{
		if (!C)
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
			if (object.Equals(A, B))
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
		}
		A = B;
		NotifyPropertyChanged(D);
		return true;
	}

	internal void A(string A, long B = 1L)
	{
		long? startMillisecs = null;
		if (this.m_A.IsRunning)
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
			startMillisecs = this.m_A.ElapsedMilliseconds;
		}
		ActionItem actionItem = new ActionItem
		{
			Desc = A,
			NumItems = Math.Max(B, 0L),
			StartMillisecs = startMillisecs,
			Status = CB.C,
			Level = checked(this.m_A.Count + 1),
			Parent = this.m_A.LastOrDefault(),
			OnRunChangeAction = this.m_A
		};
		this.m_A.Add(actionItem);
		if (this.A(actionItem))
		{
			this.B(actionItem);
		}
		E();
	}

	internal ActionItem A()
	{
		return this.m_A.LastOrDefault();
	}

	internal void A(string A)
	{
		if (this.m_A.Count == 0)
		{
			return;
		}
		checked
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
				ActionItem actionItem = this.m_A.Last();
				actionItem.CurItemNum++;
				actionItem.CurItemDesc = A;
				E();
				return;
			}
		}
	}

	internal void A(bool A)
	{
		if (this.m_A.Count == 0)
		{
			return;
		}
		ActionItem actionItem = this.m_A.Last();
		this.A(actionItem);
		if (this.m_A.Count == 1)
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
			if (this.m_A[0].ErroredOut)
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
				actionItem.Status = CB.H;
			}
			else if (A)
			{
				if (actionItem.A(CB.C, CB.E))
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
					this.A();
				}
			}
			else
			{
				actionItem.A(CB.C, CB.D);
			}
			this.m_B.Add(actionItem);
		}
		else
		{
			C(actionItem);
		}
		this.m_A.Remove(actionItem);
		E();
	}

	private void A()
	{
		ObservableCollection<ActionItem> allItems = AllItems;
		Func<ActionItem, bool> predicate;
		if (_Closure_0024__.A == null)
		{
			predicate = (_Closure_0024__.A = [SpecialName] (ActionItem A) => A.Status == CB.E);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			predicate = _Closure_0024__.A;
		}
		SkippedCount = allItems.Where(predicate).Count();
	}

	private void A(ActionItem A)
	{
		if (A.StartMillisecs.HasValue)
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
			if (this.m_A.IsRunning)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						A.DurationMillisecs = checked(this.m_A.ElapsedMilliseconds - A.StartMillisecs.Value);
						return;
					}
				}
			}
		}
		if (A.DurationMillisecs.HasValue)
		{
			A.DurationMillisecs = null;
		}
	}

	private void B()
	{
		double num;
		if (NumChecksToRun > 0)
		{
			if (CurrentCheckNum > 1)
			{
				num = (double)checked(CurrentCheckNum - 1) / (double)NumChecksToRun;
				goto IL_0047;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
		}
		num = 0.0;
		goto IL_0047;
		IL_0047:
		double num2 = num;
		ProgressArc.Percentage = num2;
		A(num2);
	}

	private void A(double A)
	{
		ShouldHidePerc = A < MinimumPercToShow;
	}

	private void C()
	{
		HeaderInfo = ((CurrentCheckNum <= NumChecksToRun) ? string.Format(VH.A(8558), CurrentCheckNum, NumChecksToRun) : string.Format(VH.A(8537), CurrentCheckNum));
	}

	internal void A(int A, bool B)
	{
		while (this.m_A.Count > A)
		{
			this.A(B);
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
			return;
		}
	}

	internal void D()
	{
		E();
	}

	private void E()
	{
		TimeSpan elapsed = this.m_A.Elapsed;
		double num = Math.Floor(elapsed.TotalHours);
		string format = VH.A(8618);
		object arg;
		if (!(num > 0.0))
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
			arg = "";
		}
		else
		{
			arg = string.Format(VH.A(7490), num);
		}
		TotalTimeStr = string.Format(format, arg, elapsed.Minutes, elapsed.Seconds);
		using (List<ActionItem>.Enumerator enumerator = this.m_A.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				ActionItem current = enumerator.Current;
				A(current);
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_00b9;
				}
				continue;
				end_IL_00b9:
				break;
			}
		}
		Action onStatusChange = OnStatusChange;
		if (onStatusChange == null)
		{
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
		onStatusChange();
	}

	internal void F()
	{
		this.m_A.Start();
	}

	internal void G()
	{
		this.m_A.Stop();
	}

	internal long A()
	{
		return this.m_A.ElapsedMilliseconds;
	}

	internal bool A(ActionItem A)
	{
		int? level = A.Level;
		bool? obj;
		if (!level.HasValue)
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
			obj = level.GetValueOrDefault() < 2;
		}
		return object.Equals(obj, true);
	}

	internal void A(RB A)
	{
		ActionItem actionItem = (A.AssociatedAction = new ActionItem
		{
			Desc = A.CheckDesc,
			Status = CB.B,
			Level = 1
		});
		if (!this.A(actionItem))
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
			AllItems.Add(actionItem);
			return;
		}
	}

	internal void B(RB A)
	{
		ActionItem associatedAction = A.AssociatedAction;
		CB? obj;
		if (associatedAction == null)
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
			obj = null;
		}
		else
		{
			obj = associatedAction.Status;
		}
		if (!object.Equals(obj, CB.B))
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
			C(associatedAction);
			return;
		}
	}

	private void B(ActionItem A)
	{
		ObservableCollection<ActionItem> allItems = AllItems;
		Func<ActionItem, bool> predicate;
		if (_Closure_0024__.B == null)
		{
			predicate = (_Closure_0024__.B = [SpecialName] (ActionItem actionItem2) => new CB[3]
			{
				CB.B,
				CB.F,
				CB.A
			}.Contains(actionItem2.Status));
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
			predicate = _Closure_0024__.B;
		}
		ActionItem actionItem = allItems.FirstOrDefault(predicate);
		int index = ((actionItem == null) ? AllItems.Count : AllItems.IndexOf(actionItem));
		AllItems.Insert(index, A);
	}

	private void C(ActionItem A)
	{
		AllItems.Remove(A);
	}

	internal bool A(string A)
	{
		return this.A([SpecialName] () => UIFormsExtensions.AskYesNo((Window)null, A, true, true));
	}

	internal void B(string A)
	{
		this.A([SpecialName] () =>
		{
			Forms.ErrorMessage(A);
			return true;
		});
	}

	internal A A<A>(Func<A> A)
	{
		bool isRunning = this.m_A.IsRunning;
		if (isRunning)
		{
			G();
		}
		try
		{
			return A();
		}
		finally
		{
			if (isRunning)
			{
				F();
			}
		}
	}
}
