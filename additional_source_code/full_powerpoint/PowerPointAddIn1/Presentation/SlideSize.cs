using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using A;

namespace PowerPointAddIn1.Presentation;

public sealed class SlideSize : INotifyPropertyChanged
{
	private string m_A;

	private double m_A;

	private long m_A;

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	public string Label
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(118726));
		}
	}

	public double BarWidth
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(118737));
		}
	}

	public long Size
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(14019));
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

	private void A(string A)
	{
		PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
		if (propertyChangedEventHandler == null)
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
			propertyChangedEventHandler(this, new PropertyChangedEventArgs(A));
			return;
		}
	}
}
