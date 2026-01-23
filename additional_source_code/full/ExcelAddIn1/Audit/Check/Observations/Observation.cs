using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Media;
using A;
using ExcelAddIn1.Audit.Check.UI;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Observations;

public class Observation : INotifyPropertyChanged
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Observation, int> A;

		public static Func<Observation, bool> A;

		public static Func<Observation, double> A;

		public static Func<Observation, string> A;

		public static Func<Observation, Severity> A;

		public static Func<Observation, string> B;

		public static Func<Observation, Severity> B;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal int A(Observation A)
		{
			return (int)A.Severity;
		}

		[SpecialName]
		internal bool A(Observation A)
		{
			return A is FC;
		}

		[SpecialName]
		internal double A(Observation A)
		{
			return A.SortIndex;
		}

		[SpecialName]
		internal string A(Observation A)
		{
			return A.GetType().FullName;
		}

		[SpecialName]
		internal Severity A(Observation A)
		{
			return A.Severity;
		}

		[SpecialName]
		internal string B(Observation A)
		{
			return A.GetType().FullName;
		}

		[SpecialName]
		internal Severity B(Observation A)
		{
			return A.Severity;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	[CompilerGenerated]
	private object m_A;

	[CompilerGenerated]
	private Worksheet m_A;

	[CompilerGenerated]
	private Range m_A;

	[CompilerGenerated]
	private Chart m_A;

	[CompilerGenerated]
	private Shape m_A;

	[CompilerGenerated]
	private string m_A;

	private string m_B;

	[CompilerGenerated]
	private string m_C;

	[CompilerGenerated]
	private Severity m_A;

	[CompilerGenerated]
	private Category m_A;

	[CompilerGenerated]
	private Observation m_A;

	[CompilerGenerated]
	private bool m_A;

	[CompilerGenerated]
	private bool m_B;

	private List<Observation> m_A;

	[CompilerGenerated]
	private string m_D;

	private int m_A;

	private int m_B;

	private int m_C;

	private bool m_C;

	private ObservableCollection<Observation> m_A;

	[CompilerGenerated]
	private bool m_D;

	private int m_D;

	private int m_E;

	private int m_F;

	private bool m_E;

	private bool m_F;

	private SolidColorBrush m_A;

	private Visibility m_A;

	private Visibility m_B;

	private FontWeight m_A;

	[CompilerGenerated]
	private Geometry m_A;

	[CompilerGenerated]
	private Thickness m_A;

	[CompilerGenerated]
	private Visibility m_C;

	[CompilerGenerated]
	private double m_A;

	[CompilerGenerated]
	private bool m_G;

	[CompilerGenerated]
	private bool H;

	[CompilerGenerated]
	private List<string> m_A;

	private bool I;

	private Visibility m_D;

	private Visibility m_E;

	private string m_E;

	private bool J;

	[CompilerGenerated]
	private bool K;

	internal object Sheet
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = RuntimeHelpers.GetObjectValue(value);
		}
	}

	internal Worksheet Worksheet
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

	internal Range Range
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

	internal Chart Chart
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

	internal Shape Shape
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

	public int SheetIndex
	{
		get
		{
			if (Worksheet != null)
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
						return Worksheet.Index;
					}
				}
			}
			if (Chart != null)
			{
				return Chart.Index;
			}
			if (Sheet is Worksheet)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						return ((Worksheet)Sheet).Index;
					}
				}
			}
			if (Sheet is Chart)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						return ((Chart)Sheet).Index;
					}
				}
			}
			return -1;
		}
	}

	public string SheetName
	{
		get
		{
			if (Worksheet != null)
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
						return Worksheet.Name;
					}
				}
			}
			if (Chart != null)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						return Chart.Name;
					}
				}
			}
			if (Sheet is Worksheet)
			{
				return ((Worksheet)Sheet).Name;
			}
			if (Sheet is Chart)
			{
				return ((Chart)Sheet).Name;
			}
			return VH.A(21448);
		}
	}

	public string Title
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

	public string Subtitle
	{
		get
		{
			return this.m_B;
		}
		set
		{
			B(ref this.m_B, value, C: false, VH.A(21471));
		}
	}

	public string Explanation
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

	public Severity Severity
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

	internal Category Category
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

	internal Observation Parent
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
			return this.m_A;
		}
	}

	public bool IsExcessObs
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
	}

	internal virtual List<Observation> Children
	{
		get
		{
			return this.m_A;
		}
		set
		{
			List<Observation> list = value;
			if (list == null)
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
				list = new List<Observation>();
			}
			this.m_A = list;
			this.m_A.ForEach([SpecialName] (Observation A) =>
			{
				A.Parent = this;
			});
		}
	}

	public string GroupName
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
	}

	public int ErrorsGroupCount
	{
		get
		{
			return this.m_A;
		}
		set
		{
			B(ref this.m_A, value, C: false, VH.A(21488));
		}
	}

	public int WarningsGroupCount
	{
		get
		{
			return this.m_B;
		}
		set
		{
			B(ref this.m_B, value, C: false, VH.A(21521));
		}
	}

	public int MessagesGroupCount
	{
		get
		{
			return this.m_C;
		}
		set
		{
			B(ref this.m_C, value, C: false, VH.A(21558));
		}
	}

	public bool IsExpanded
	{
		get
		{
			return this.m_C;
		}
		set
		{
			if (!B(ref this.m_C, value, C: false, VH.A(21595)) || AllItems == null)
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
				B();
				return;
			}
		}
	}

	internal ObservableCollection<Observation> AllItems
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
				switch (2)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.m_A = value;
				C();
				return;
			}
		}
	}

	public bool IsAction
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
	}

	public int ErrorsCount
	{
		get
		{
			return this.m_D;
		}
		set
		{
			this.m_D = value;
			NotifyPropertyChanged(VH.A(21616));
		}
	}

	public int WarningsCount
	{
		get
		{
			return this.m_E;
		}
		set
		{
			this.m_E = value;
			NotifyPropertyChanged(VH.A(21639));
		}
	}

	public int MessagesCount
	{
		get
		{
			return this.m_F;
		}
		set
		{
			this.m_F = value;
			NotifyPropertyChanged(VH.A(21666));
		}
	}

	internal bool IsHidden
	{
		get
		{
			return this.m_E;
		}
		set
		{
			this.m_E = value;
			int visibility;
			if (!value)
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
				visibility = 2;
			}
			else
			{
				visibility = 0;
			}
			Visibility = (Visibility)visibility;
		}
	}

	public bool IsSelected
	{
		get
		{
			return this.m_F;
		}
		set
		{
			this.m_F = value;
			NotifyPropertyChanged(VH.A(21693));
			if (value)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
					{
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						object obj = ColorConverter.ConvertFromString(VH.A(21714));
						Color color;
						if (obj == null)
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
							color = default(Color);
						}
						else
						{
							color = (Color)obj;
						}
						TypeColor = new SolidColorBrush(color);
						return;
					}
					}
				}
			}
			E();
		}
	}

	public SolidColorBrush TypeColor
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			NotifyPropertyChanged(VH.A(21729));
		}
	}

	public Visibility SubtitleVisibility
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			NotifyPropertyChanged(VH.A(21748));
		}
	}

	public Visibility Visibility
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			NotifyPropertyChanged(VH.A(21785));
		}
	}

	public FontWeight FontWeight
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			NotifyPropertyChanged(VH.A(21806));
		}
	}

	public Geometry Icon
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

	public Thickness IconPadding
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

	public Visibility TooltipVisibility
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

	public double SortIndex
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

	public bool HasFix
	{
		[CompilerGenerated]
		get
		{
			return this.m_G;
		}
		[CompilerGenerated]
		set
		{
			this.m_G = value;
		}
	}

	public bool CanFixMultiple
	{
		[CompilerGenerated]
		get
		{
			return H;
		}
		[CompilerGenerated]
		set
		{
			H = value;
		}
	}

	public List<string> DisplayText
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

	public bool IsFixed
	{
		get
		{
			return I;
		}
		set
		{
			I = value;
			if (!value)
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
				E();
				FixIconPath = Constants.FIX_ICON_PATH_WRENCH;
				FixEnabled = true;
			}
			else
			{
				object obj = ColorConverter.ConvertFromString(VH.A(21714));
				Color color;
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
					color = default(Color);
				}
				else
				{
					color = (Color)obj;
				}
				TypeColor = new SolidColorBrush(color);
				FixIconPath = VH.A(21827);
				FixEnabled = false;
			}
			G();
			NotifyPropertyChanged(VH.A(21938));
		}
	}

	public Visibility FixVisibility
	{
		get
		{
			return this.m_D;
		}
		set
		{
			this.m_D = value;
			NotifyPropertyChanged(VH.A(21953));
		}
	}

	public Visibility FixMenuVisibility
	{
		get
		{
			return this.m_E;
		}
		set
		{
			this.m_E = value;
			NotifyPropertyChanged(VH.A(21980));
		}
	}

	public string FixIconPath
	{
		get
		{
			return this.m_E;
		}
		set
		{
			this.m_E = value;
			NotifyPropertyChanged(VH.A(22015));
		}
	}

	public bool FixEnabled
	{
		get
		{
			return J;
		}
		set
		{
			J = value;
			NotifyPropertyChanged(VH.A(22038));
		}
	}

	internal bool AffectsGroupCount
	{
		[CompilerGenerated]
		get
		{
			return K;
		}
		[CompilerGenerated]
		set
		{
			K = value;
		}
	}

	private event PropertyChangedEventHandler PropertyChanged
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
				switch (7)
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

	protected Observation()
	{
		Sheet = null;
		Worksheet = null;
		Range = null;
		Chart = null;
		Shape = null;
		this.m_B = this is FC;
		this.m_A = new List<Observation>();
		this.m_D = false;
		this.m_F = false;
		HasFix = false;
		CanFixMultiple = false;
		I = false;
		J = true;
		AffectsGroupCount = true;
	}

	public Observation(Category cat, Severity sev, string strTitle)
	{
		Sheet = null;
		Worksheet = null;
		Range = null;
		Chart = null;
		Shape = null;
		this.m_B = this is FC;
		this.m_A = new List<Observation>();
		this.m_D = false;
		this.m_F = false;
		HasFix = false;
		CanFixMultiple = false;
		I = false;
		J = true;
		AffectsGroupCount = true;
		Category = cat;
		Severity = sev;
		Title = strTitle;
	}

	public Observation(Category cat, Severity sev, string strTitle, Worksheet ws)
	{
		Sheet = null;
		Worksheet = null;
		Range = null;
		Chart = null;
		Shape = null;
		this.m_B = this is FC;
		this.m_A = new List<Observation>();
		this.m_D = false;
		this.m_F = false;
		HasFix = false;
		CanFixMultiple = false;
		I = false;
		J = true;
		AffectsGroupCount = true;
		Category = cat;
		Severity = sev;
		Title = strTitle;
		Sheet = ws;
		Worksheet = ws;
	}

	public Observation(Category cat, Severity sev, string strTitle, Range rng)
	{
		Sheet = null;
		Worksheet = null;
		Range = null;
		Chart = null;
		Shape = null;
		this.m_B = this is FC;
		this.m_A = new List<Observation>();
		this.m_D = false;
		this.m_F = false;
		HasFix = false;
		CanFixMultiple = false;
		I = false;
		J = true;
		AffectsGroupCount = true;
		Category = cat;
		Severity = sev;
		Title = strTitle;
		Subtitle = VH.A(11531) + rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		Sheet = rng.Worksheet;
		Worksheet = rng.Worksheet;
		Range = rng;
	}

	public Observation(Category cat, Severity sev, string strTitle, Chart cht)
	{
		Sheet = null;
		Worksheet = null;
		Range = null;
		Chart = null;
		Shape = null;
		this.m_B = this is FC;
		this.m_A = new List<Observation>();
		this.m_D = false;
		this.m_F = false;
		HasFix = false;
		CanFixMultiple = false;
		I = false;
		J = true;
		AffectsGroupCount = true;
		Category = cat;
		Severity = sev;
		Title = strTitle;
		Sheet = cht;
		Chart = cht;
	}

	private Observation(List<Observation> A, int B)
	{
		Sheet = null;
		Worksheet = null;
		Range = null;
		Chart = null;
		Shape = null;
		this.m_B = this is FC;
		this.m_A = new List<Observation>();
		this.m_D = false;
		this.m_F = false;
		HasFix = false;
		CanFixMultiple = false;
		I = false;
		J = true;
		AffectsGroupCount = true;
		Category = Category.Group;
		this.m_A = true;
		Children = A;
		Observation observation = Children.FirstOrDefault();
		object title;
		if (observation == null)
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
			title = null;
		}
		else
		{
			title = observation.Title;
		}
		Title = (string)title;
		object explanation;
		if (observation == null)
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
			explanation = null;
		}
		else
		{
			explanation = observation.Explanation;
		}
		Explanation = (string)explanation;
		object d;
		if (observation == null)
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
			d = null;
		}
		else
		{
			d = observation.Title;
		}
		this.m_D = (string)d;
		int value;
		checked
		{
			value = (int)Math.Round(Math.Round(Children.Average([SpecialName] (Observation observation2) => unchecked((int)observation2.Severity)), 0));
		}
		Severity = (Severity)Conversions.ToInteger(Enum.ToObject(typeof(Severity), value));
		ErrorsGroupCount = global::A.GC.A(Children);
		WarningsGroupCount = global::A.GC.B(Children);
		MessagesGroupCount = global::A.GC.C(Children);
		Children = Observation.B(A, B);
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

	private bool B<A>(ref A A, A B, bool C = false, [CallerMemberName] string D = null)
	{
		if (!C)
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
			if (object.Equals(A, B))
			{
				return false;
			}
		}
		A = B;
		NotifyPropertyChanged(D);
		return true;
	}

	private void B(Severity A, int B)
	{
		checked
		{
			if (A != Severity.Medium)
			{
				if (A == Severity.High)
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
					ErrorsGroupCount += B;
				}
				else
				{
					MessagesGroupCount += B;
				}
			}
			else
			{
				WarningsGroupCount += B;
			}
			Observation parent = Parent;
			if (parent == null)
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
			parent.B(A, B);
		}
	}

	private void B()
	{
		if (Children.Count == 0)
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
			int num = AllItems.IndexOf(this);
			if (num < 0)
			{
				return;
			}
			int num2 = num;
			using List<Observation>.Enumerator enumerator = Children.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Observation current = enumerator.Current;
				if (this.m_C)
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
					AllItems.Insert(num2, current);
					num2 = checked(num2 + 1);
				}
				else
				{
					AllItems.Remove(current);
				}
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

	private void C()
	{
		using List<Observation>.Enumerator enumerator = Children.GetEnumerator();
		while (enumerator.MoveNext())
		{
			Observation current = enumerator.Current;
			current.AllItems = AllItems;
			current.C();
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
			return;
		}
	}

	internal static Observation B()
	{
		return new Observation();
	}

	internal static Observation B(List<Observation> A, int B)
	{
		return new Observation(A, B);
	}

	public virtual void FixAction()
	{
	}

	public virtual void FixAction(int index)
	{
	}

	internal void B(int A, int B, int C, ref double D)
	{
		D += 1.0;
		ErrorsCount = A;
		WarningsCount = B;
		MessagesCount = C;
		SortIndex = D;
		F();
		E();
		int subtitleVisibility;
		if (Subtitle != null)
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
			if (Subtitle.Length > 0)
			{
				subtitleVisibility = 0;
				goto IL_006e;
			}
		}
		subtitleVisibility = 2;
		goto IL_006e;
		IL_006e:
		SubtitleVisibility = (Visibility)subtitleVisibility;
		int tooltipVisibility;
		if (Explanation != null)
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
			if (Explanation.Length > 0)
			{
				tooltipVisibility = 0;
				goto IL_00a4;
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
		tooltipVisibility = 2;
		goto IL_00a4;
		IL_00a4:
		TooltipVisibility = (Visibility)tooltipVisibility;
		FontWeight = FontWeights.SemiBold;
		IsHidden = false;
		IsFixed = false;
	}

	internal virtual bool A(Observation A)
	{
		if (!Children.Remove(A))
		{
			return false;
		}
		ObservableCollection<Observation> allItems = AllItems;
		if (allItems == null)
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
		}
		else
		{
			allItems.Remove(A);
		}
		A.Parent = null;
		if (A.AffectsGroupCount)
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
			B(A.Severity, -1);
		}
		Observation observation = Children.FirstOrDefault([SpecialName] (Observation observation2) => observation2 is FC);
		if (observation != null)
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
			observation.D();
		}
		if (!(this is FC) && Parent == null)
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
			if (Children.Count < 1)
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
				ObservableCollection<Observation> allItems2 = AllItems;
				if (allItems2 == null)
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
				}
				else
				{
					allItems2.Remove(this);
				}
			}
		}
		return true;
	}

	private void D()
	{
		if (Children.Count > 0)
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
			if (Parent != null)
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
				Observation a = Children[0];
				A(a);
				B(a);
			}
		}
		if (!(this is FC))
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
			if (Parent == null)
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
				if (Children.Count < 1)
				{
					Parent.A(this);
					return;
				}
				if (Children.Count != 1)
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
					D();
					return;
				}
			}
		}
	}

	private void B(Observation A)
	{
		Parent.Children.Insert(Parent.Children.IndexOf(this), A);
		A.Parent = Parent;
		if (A.AffectsGroupCount)
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
			Parent.B(A.Severity, 1);
		}
		int? num = AllItems?.IndexOf(this);
		int? num2 = num;
		if ((num2.HasValue ? new bool?(num2.GetValueOrDefault() > -1) : ((bool?)null)) != true)
		{
			return;
		}
		IEnumerable<Observation> source = AllItems.Where([SpecialName] (Observation observation) => observation.SortIndex < SortIndex);
		double num3;
		if (!source.Any())
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
			num3 = 0.0;
		}
		else
		{
			Func<Observation, double> selector;
			if (_Closure_0024__.A == null)
			{
				selector = (_Closure_0024__.A = [SpecialName] (Observation observation) => observation.SortIndex);
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
				selector = _Closure_0024__.A;
			}
			num3 = source.Max(selector);
		}
		double num4 = num3;
		A.SortIndex = (SortIndex + num4) / 2.0;
		AllItems.Insert(num.Value, A);
	}

	private void E()
	{
		Color color = default(Color);
		switch (Severity)
		{
		case Severity.Medium:
			color = Constants.COLOR_YELLOW;
			break;
		case Severity.High:
			color = Constants.COLOR_RED;
			break;
		case Severity.Low:
			color = Constants.COLOR_BLUE;
			break;
		}
		TypeColor = new SolidColorBrush(color);
	}

	private void F()
	{
		IconPadding = new Thickness(2.0);
		IconCache cache = Icons.Cache;
		switch (Category)
		{
		case Category.FormulaErrors:
		case Category.FormulaComplexity:
		case Category.FormulaIntegrity:
			Icon = cache.GeoFormula;
			break;
		case Category.BrandCompliance:
			Icon = cache.GeoPalette;
			IconPadding = new Thickness(1.0);
			break;
		case Category.HiddenData:
			Icon = cache.GeoEye;
			break;
		case Category.Data:
			Icon = cache.GeoData;
			break;
		case Category.ModelStructure:
			Icon = cache.GeoPuzzle;
			IconPadding = new Thickness(3.0);
			break;
		case Category.PrivacySecurity:
			Icon = cache.GeoShield;
			break;
		case Category.Workbook:
			Icon = cache.GeoFile;
			break;
		case Category.Performance:
			Icon = cache.GeoSpeed;
			break;
		case Category.BestPractices:
			Icon = cache.GeoMedal;
			IconPadding = new Thickness(3.0);
			break;
		case Category.Oddities:
			Icon = cache.GeoQuestion;
			break;
		default:
			Icon = cache.GeoQuestion;
			break;
		}
		cache = null;
		Icon.Freeze();
	}

	private void G()
	{
		int num;
		if (!IsGrouper)
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
			if (!IsFixed && !HasFix)
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
				List<string> displayText = DisplayText;
				bool? obj;
				if (displayText == null)
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
					obj = null;
				}
				else
				{
					obj = displayText.Count > 1;
				}
				num = (object.Equals(obj, true) ? 1 : 0);
			}
			else
			{
				num = 1;
			}
		}
		else
		{
			num = 0;
		}
		bool flag = (byte)num != 0;
		int num2;
		if (!IsGrouper)
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
			if (!IsFixed)
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
				List<string> displayText2 = DisplayText;
				num2 = (object.Equals((displayText2 != null) ? new bool?(displayText2.Count > 1) : ((bool?)null), true) ? 1 : 0);
				goto IL_00e2;
			}
		}
		num2 = 0;
		goto IL_00e2;
		IL_00e2:
		bool flag2 = (byte)num2 != 0;
		int fixVisibility;
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
			fixVisibility = 2;
		}
		else
		{
			fixVisibility = 0;
		}
		FixVisibility = (Visibility)fixVisibility;
		int fixMenuVisibility;
		if (!flag2)
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
			fixMenuVisibility = 2;
		}
		else
		{
			fixMenuVisibility = 0;
		}
		FixMenuVisibility = (Visibility)fixMenuVisibility;
	}

	private static List<Observation> B(List<Observation> A, int B)
	{
		if (A.Count <= B)
		{
			return A;
		}
		bool flag = false;
		IEnumerator<IGrouping<string, Observation>> enumerator = default(IEnumerator<IGrouping<string, Observation>>);
		try
		{
			List<Observation> source = A;
			Func<Observation, string> keySelector;
			if (_Closure_0024__.A == null)
			{
				keySelector = (_Closure_0024__.A = [SpecialName] (Observation observation) => observation.GetType().FullName);
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
				keySelector = _Closure_0024__.A;
			}
			enumerator = source.GroupBy(keySelector).GetEnumerator();
			IEnumerator<IGrouping<Severity, Observation>> enumerator2 = default(IEnumerator<IGrouping<Severity, Observation>>);
			while (enumerator.MoveNext())
			{
				IGrouping<string, Observation> current = enumerator.Current;
				try
				{
					Func<Observation, Severity> keySelector2;
					if (_Closure_0024__.A == null)
					{
						keySelector2 = (_Closure_0024__.A = [SpecialName] (Observation observation) => observation.Severity);
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
						keySelector2 = _Closure_0024__.A;
					}
					enumerator2 = current.GroupBy(keySelector2).GetEnumerator();
					while (enumerator2.MoveNext())
					{
						if (enumerator2.Current.Count() <= B)
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
							flag = true;
							break;
						}
						break;
					}
				}
				finally
				{
					enumerator2?.Dispose();
				}
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_00ed;
				}
				continue;
				end_IL_00ed:
				break;
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
		if (!flag)
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
		List<Observation> list = new List<Observation>();
		if (B < 1)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					return list;
				}
			}
		}
		checked
		{
			IEnumerator<IGrouping<string, Observation>> enumerator3 = default(IEnumerator<IGrouping<string, Observation>>);
			try
			{
				List<Observation> source2 = A;
				Func<Observation, string> keySelector3;
				if (_Closure_0024__.B == null)
				{
					keySelector3 = (_Closure_0024__.B = [SpecialName] (Observation observation) => observation.GetType().FullName);
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
					keySelector3 = _Closure_0024__.B;
				}
				enumerator3 = source2.GroupBy(keySelector3).GetEnumerator();
				IEnumerator<IGrouping<Severity, Observation>> enumerator4 = default(IEnumerator<IGrouping<Severity, Observation>>);
				while (enumerator3.MoveNext())
				{
					IGrouping<string, Observation> current2 = enumerator3.Current;
					try
					{
						Func<Observation, Severity> keySelector4;
						if (_Closure_0024__.B == null)
						{
							keySelector4 = (_Closure_0024__.B = [SpecialName] (Observation observation) => observation.Severity);
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
							keySelector4 = _Closure_0024__.B;
						}
						enumerator4 = current2.GroupBy(keySelector4).GetEnumerator();
						while (enumerator4.MoveNext())
						{
							IGrouping<Severity, Observation> current3 = enumerator4.Current;
							if (current3.Count() <= B)
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
								list.AddRange(current3);
							}
							else
							{
								list.AddRange(current3.Take(B - 1));
								list.Add(new FC(current3.Skip(B - 1)));
							}
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_0227;
							}
							continue;
							end_IL_0227:
							break;
						}
					}
					finally
					{
						enumerator4?.Dispose();
					}
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_024b;
					}
					continue;
					end_IL_024b:
					break;
				}
			}
			finally
			{
				if (enumerator3 != null)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						enumerator3.Dispose();
						break;
					}
				}
			}
			A = null;
			return list;
		}
	}

	[SpecialName]
	[CompilerGenerated]
	private void C(Observation A)
	{
		A.Parent = this;
	}

	[SpecialName]
	[CompilerGenerated]
	private bool B(Observation A)
	{
		return A.SortIndex < SortIndex;
	}
}
