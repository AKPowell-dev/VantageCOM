using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using A;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Model;

public sealed class Statistic : INotifyPropertyChanged
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private string m_A;

	private string B;

	[CompilerGenerated]
	private string C;

	[CompilerGenerated]
	private string D;

	public bool IsChecked
	{
		get
		{
			return Conversions.ToBoolean(this.m_A);
		}
		set
		{
			this.m_A = Conversions.ToString(value);
			A(VH.A(90018));
		}
	}

	public string Label
	{
		get
		{
			return B;
		}
		set
		{
			B = value;
			A(VH.A(49999));
		}
	}

	public string FunctionEnglish
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	public string FunctionLocal
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[CompilerGenerated]
		set
		{
			D = value;
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

	public Statistic(bool blnChecked, string strLabel, string strFxnEnglish, string strFxnLocal)
	{
		IsChecked = blnChecked;
		Label = strLabel;
		FunctionEnglish = strFxnEnglish;
		FunctionLocal = strFxnLocal;
	}

	private void A(string A)
	{
		this.m_A?.Invoke(this, new PropertyChangedEventArgs(A));
	}
}
