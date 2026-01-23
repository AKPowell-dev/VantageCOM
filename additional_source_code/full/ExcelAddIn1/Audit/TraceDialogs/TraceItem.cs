using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Media;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.TraceDialogs;

public class TraceItem : INotifyPropertyChanged
{
	[CompilerGenerated]
	private PropertyChangedEventHandler A;

	private bool A;

	private bool B;

	private Geometry A;

	private string A;

	[CompilerGenerated]
	private int A;

	[CompilerGenerated]
	private TraceItem A;

	private RangeObservableCollection<TraceItem> A;

	[CompilerGenerated]
	private Range A;

	[CompilerGenerated]
	private int B;

	[CompilerGenerated]
	private List<Range> A;

	public bool IsSelected
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			NotifyPropertyChanged(VH.A(21693));
		}
	}

	public bool IsExpanded
	{
		get
		{
			return this.B;
		}
		set
		{
			this.B = value;
			NotifyPropertyChanged(VH.A(21595));
		}
	}

	public Geometry Icon
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			NotifyPropertyChanged(VH.A(49990));
		}
	}

	public string Label
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			NotifyPropertyChanged(VH.A(49999));
		}
	}

	public int Level
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public TraceItem Parent
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public RangeObservableCollection<TraceItem> Items
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			NotifyPropertyChanged(VH.A(50010));
		}
	}

	public Range Range
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public int Index
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public List<Range> ExtraAreas
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	private event PropertyChangedEventHandler PropertyChanged
	{
		[CompilerGenerated]
		add
		{
			PropertyChangedEventHandler propertyChangedEventHandler = this.A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Combine(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.A, value2, propertyChangedEventHandler2);
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
		[CompilerGenerated]
		remove
		{
			PropertyChangedEventHandler propertyChangedEventHandler = this.A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Remove(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.A, value2, propertyChangedEventHandler2);
			}
			while ((object)propertyChangedEventHandler != propertyChangedEventHandler2);
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
	}

	public TraceItem(TraceItem p, Range rng, int intLevel, string strIcon)
	{
		Parent = p;
		Range = rng;
		Level = intLevel;
		Items = new RangeObservableCollection<TraceItem>();
		if (strIcon.Length <= 0)
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
			Icon = Geometry.Parse(strIcon);
			Icon.Freeze();
			return;
		}
	}

	public void NotifyPropertyChanged(string propertyName)
	{
		PropertyChangedEventHandler propertyChangedEventHandler = this.A;
		if (propertyChangedEventHandler == null)
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
			propertyChangedEventHandler(this, new PropertyChangedEventArgs(propertyName));
			return;
		}
	}
}
