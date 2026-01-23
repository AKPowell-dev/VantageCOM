using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using A;
using MacabacusMacros;
using MacabacusMacros.Xaml.Geometry;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Helpers;

public sealed class ActionItem : INotifyPropertyChanged
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private Arc m_A;

	private string m_A;

	private string m_B;

	private string m_C;

	private long m_A;

	private long m_B;

	private string m_D;

	[CompilerGenerated]
	private double m_A;

	private bool m_A;

	[CompilerGenerated]
	private long? m_A;

	private long? m_B;

	private string m_E;

	private int? m_A;

	private ActionItem m_A;

	private bool m_B;

	private ActionItem m_B;

	private readonly object m_A;

	private CB m_A;

	private bool m_C;

	private bool m_D;

	private bool m_E;

	private bool m_F;

	private bool m_G;

	private bool m_H;

	private bool m_I;

	private bool m_J;

	private string m_F;

	private Exception m_A;

	[CompilerGenerated]
	private Action<List<ActionItem>> m_A;

	[CompilerGenerated]
	private bool m_K;

	[CompilerGenerated]
	private bool m_L;

	[CompilerGenerated]
	private bool m_M;

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

	public string IndentStr
	{
		get
		{
			return this.m_A;
		}
		set
		{
			A(ref this.m_A, value, C: false, VH.A(7297));
		}
	}

	public string Desc
	{
		get
		{
			return this.m_B;
		}
		set
		{
			A(ref this.m_B, value, C: false, VH.A(7316));
		}
	}

	internal string CurItemDesc
	{
		get
		{
			return this.m_C;
		}
		set
		{
			if (!A(ref this.m_C, value, C: false, VH.A(7325)))
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
				C();
				return;
			}
		}
	}

	internal long CurItemNum
	{
		get
		{
			return this.m_A;
		}
		set
		{
			if (!A(ref this.m_A, value, C: false, VH.A(7348)))
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
				D();
				B();
				return;
			}
		}
	}

	internal long NumItems
	{
		get
		{
			return this.m_B;
		}
		set
		{
			if (!A(ref this.m_B, value, C: false, VH.A(7369)))
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
				D();
				return;
			}
		}
	}

	public string ItemsStr
	{
		get
		{
			return this.m_D;
		}
		set
		{
			A(ref this.m_D, value, C: false, VH.A(7386));
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

	internal long? StartMillisecs
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

	internal long? DurationMillisecs
	{
		get
		{
			return this.m_B;
		}
		set
		{
			if (!A(ref this.m_B, value, C: false, VH.A(7432)))
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
				E();
				return;
			}
		}
	}

	public string DurationStr
	{
		get
		{
			return this.m_E;
		}
		set
		{
			A(ref this.m_E, value, C: false, VH.A(7467));
		}
	}

	internal int? Level
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
				switch (4)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.m_A = value;
				A();
				K();
				return;
			}
		}
	}

	public ActionItem Child
	{
		get
		{
			return this.m_A;
		}
		private set
		{
			if (!A(ref this.m_A, value, C: false, VH.A(7532)))
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
				F();
				return;
			}
		}
	}

	public bool HasChild
	{
		get
		{
			return this.m_B;
		}
		set
		{
			A(ref this.m_B, value, C: false, VH.A(7543));
		}
	}

	internal ActionItem Parent
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
				switch (4)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.m_B = value;
				if (this.m_B == null)
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
					this.m_B.Child = this;
					return;
				}
			}
		}
	}

	internal CB Status
	{
		get
		{
			return this.m_A;
		}
		set
		{
			object a = this.m_A;
			ObjectFlowControl.CheckForSyncLockOnValueType(a);
			bool lockTaken = false;
			try
			{
				Monitor.Enter(a, ref lockTaken);
				if (!A(ref this.m_A, value, C: false, VH.A(7560)))
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
							return;
						}
					}
				}
				G();
				H();
				I();
				J();
				L();
				M();
			}
			finally
			{
				if (lockTaken)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						Monitor.Exit(a);
						break;
					}
				}
			}
		}
	}

	public bool IsPending
	{
		get
		{
			return this.m_C;
		}
		set
		{
			A(ref this.m_C, value, C: false, VH.A(7573));
		}
	}

	public bool IsRunning
	{
		get
		{
			return this.m_D;
		}
		set
		{
			if (!A(ref this.m_D, value, C: false, VH.A(7592)))
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
				K();
				if (this.m_D)
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
					ActionItem child = Child;
					bool? obj;
					if (child == null)
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
					else
					{
						obj = child.IsRunning;
					}
					if (!object.Equals(obj, true))
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
						Child.Status = Status;
						return;
					}
				}
			}
		}
	}

	public bool IsCompletedFully
	{
		get
		{
			return this.m_E;
		}
		set
		{
			A(ref this.m_E, value, C: false, VH.A(7611));
		}
	}

	public bool IsSkipped
	{
		get
		{
			return this.m_F;
		}
		set
		{
			A(ref this.m_F, value, C: false, VH.A(7644));
		}
	}

	public bool IsSkippable
	{
		get
		{
			return this.m_G;
		}
		set
		{
			A(ref this.m_G, value, C: false, VH.A(7663));
		}
	}

	public bool IsToBeBypassed
	{
		get
		{
			return this.m_H;
		}
		set
		{
			A(ref this.m_H, value, C: false, VH.A(7686));
		}
	}

	public bool IsBypassed
	{
		get
		{
			return this.m_I;
		}
		set
		{
			A(ref this.m_I, value, C: false, VH.A(7715));
		}
	}

	public bool ErroredOut
	{
		get
		{
			return this.m_J;
		}
		set
		{
			A(ref this.m_J, value, C: false, VH.A(7736));
		}
	}

	public string ExcMessage
	{
		get
		{
			return this.m_F;
		}
		set
		{
			A(ref this.m_F, value, C: false, VH.A(7757));
		}
	}

	internal Exception Exception
	{
		get
		{
			return this.m_A;
		}
		set
		{
			if (A(ref this.m_A, value, C: false, VH.A(7808)))
			{
				N();
				O();
			}
		}
	}

	internal Action<List<ActionItem>> OnRunChangeAction
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

	public bool IsGrouper
	{
		[CompilerGenerated]
		get
		{
			return this.m_K;
		}
	}

	public bool IsAction
	{
		[CompilerGenerated]
		get
		{
			return this.m_L;
		}
	}

	public bool IsExcessObs
	{
		[CompilerGenerated]
		get
		{
			return this.m_M;
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

	internal ActionItem()
	{
		//IL_0081: Unknown result type (might be due to invalid IL or missing references)
		//IL_008b: Expected O, but got Unknown
		this.m_A = "";
		MinimumPercToShow = 0.1;
		this.m_E = string.Empty;
		this.m_A = RuntimeHelpers.GetObjectValue(new object());
		this.m_A = CB.A;
		this.m_F = "";
		this.m_K = false;
		this.m_L = true;
		this.m_M = false;
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
			switch (2)
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
		if (!C && object.Equals(A, B))
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
					return false;
				}
			}
		}
		A = B;
		NotifyPropertyChanged(D);
		return true;
	}

	private void A()
	{
		IndentStr = Strings.Space(checked(((Level ?? 1) - 1) * 4));
	}

	private void B()
	{
		List<ActionItem> list = new List<ActionItem> { this };
		for (ActionItem parent = Parent; parent != null; parent = parent.Parent)
		{
			list.Add(parent);
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
			Action<List<ActionItem>> onRunChangeAction = OnRunChangeAction;
			if (onRunChangeAction == null)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
			onRunChangeAction(list);
			return;
		}
	}

	private void C()
	{
		ItemsStr = modFunctionsStr.BlankTo(CurItemDesc, "");
	}

	private void D()
	{
		double num;
		if (NumItems > 0)
		{
			if (CurItemNum > 1)
			{
				num = (double)checked(CurItemNum - 1) / (double)NumItems;
				goto IL_004c;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
		}
		num = 0.0;
		goto IL_004c;
		IL_004c:
		double num2 = num;
		ProgressArc.Percentage = num2;
		A(num2);
	}

	private void A(double A)
	{
		ShouldHidePerc = A < MinimumPercToShow;
	}

	private void E()
	{
		object durationStr;
		if (DurationMillisecs.HasValue)
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
			durationStr = A(DurationMillisecs.Value);
		}
		else
		{
			durationStr = "";
		}
		DurationStr = (string)durationStr;
	}

	private static string A(long A)
	{
		double num = Math.Round((double)A / 1000.0, 0);
		double num2 = Math.Floor(num / 60.0);
		num -= num2 * 60.0;
		double num3 = Math.Floor(num2 / 60.0);
		num2 -= num3 * 60.0;
		StringBuilder stringBuilder = new StringBuilder();
		checked
		{
			if (num3 > 0.0)
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
				stringBuilder.Append(string.Format(VH.A(7490), (long)Math.Round(num3)));
			}
			stringBuilder.Append(string.Format(VH.A(7505), (long)Math.Round(num2), (long)Math.Round(num)));
			return stringBuilder.ToString();
		}
	}

	private void F()
	{
		HasChild = Child != null;
	}

	internal bool A(CB A, CB B)
	{
		object a = this.m_A;
		ObjectFlowControl.CheckForSyncLockOnValueType(a);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(a, ref lockTaken);
			if (Status != A)
			{
				return false;
			}
			Status = B;
			return true;
		}
		finally
		{
			if (lockTaken)
			{
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
					Monitor.Exit(a);
					break;
				}
			}
		}
	}

	private void G()
	{
		IsPending = Status == CB.B;
	}

	private void H()
	{
		IsRunning = Status == CB.C;
	}

	private void I()
	{
		IsCompletedFully = Status == CB.D;
	}

	private void J()
	{
		IsSkipped = Status == CB.E;
	}

	private void K()
	{
		IsSkippable = object.Equals(Level, 1) && IsRunning;
	}

	private void L()
	{
		IsToBeBypassed = Status == CB.F;
	}

	private void M()
	{
		IsBypassed = Status == CB.G;
	}

	internal bool B()
	{
		if (!IsBypassed && !IsToBeBypassed)
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
					return IsPending;
				}
			}
		}
		return true;
	}

	private void N()
	{
		ErroredOut = Exception != null;
	}

	private void O()
	{
		object excMessage;
		if (Exception != null)
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
			excMessage = string.Format(VH.A(7778), VH.A(7803), Exception.Message);
		}
		else
		{
			excMessage = "";
		}
		ExcMessage = (string)excMessage;
	}
}
