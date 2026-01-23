using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Workbook;

public sealed class DefinedName : INotifyPropertyChanged
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private Name m_A;

	private int m_A;

	private bool m_A;

	private string m_A;

	private string B;

	private string C;

	private string D;

	private float m_A;

	public Name Name
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(19019));
		}
	}

	public int Index
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(48135));
		}
	}

	public bool IsChecked
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(90018));
		}
	}

	public string Label
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(49999));
		}
	}

	public string Text
	{
		get
		{
			return B;
		}
		set
		{
			B = value;
			A(VH.A(96399));
		}
	}

	public string ParentName
	{
		get
		{
			return C;
		}
		set
		{
			C = value;
			A(VH.A(176850));
		}
	}

	public string RefersTo
	{
		get
		{
			return D;
		}
		set
		{
			D = value;
			A(VH.A(153696));
		}
	}

	public float Opacity
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(123854));
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

	public DefinedName(Name nm, int intIndex, bool blnChecked, string strLabel, string strText, string strParentName, string strRefersTo, float sngOpacity)
	{
		Name = nm;
		IsChecked = blnChecked;
		Label = strLabel;
		Text = strText;
		ParentName = strParentName;
		RefersTo = strRefersTo;
		Opacity = sngOpacity;
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
			switch (2)
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
