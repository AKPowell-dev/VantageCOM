using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using A;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.DocBuilder;

public class BaseQuestion : INotifyPropertyChanged
{
	[CompilerGenerated]
	private PropertyChangedEventHandler A;

	private ContentControl A;

	private string A;

	private Visibility A;

	public ContentControl ContentControl
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public string Question
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public Visibility ApplyButtonVisibility
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
			NotifyPropertyChanged(XC.A(21319));
		}
	}

	public event PropertyChangedEventHandler PropertyChanged
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
			PropertyChangedEventHandler propertyChangedEventHandler = this.A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Remove(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.A, value2, propertyChangedEventHandler2);
			}
			while ((object)propertyChangedEventHandler != propertyChangedEventHandler2);
		}
	}

	public BaseQuestion(ContentControl cc, int intIndex)
	{
		Question = intIndex + XC.A(21362) + cc.Title;
		ContentControl = cc;
		ApplyButtonVisibility = Visibility.Hidden;
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
			switch (6)
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
