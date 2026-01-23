using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.Shapes.Arrange;

public abstract class Arrangement : INotifyPropertyChanged
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private bool m_A;

	private int m_A;

	private Visibility m_A;

	public bool IsChecked
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			C(AH.A(12198));
		}
	}

	public int SlideCount
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			C(AH.A(63364));
			int adornerVisibility;
			if (value != 1)
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
				adornerVisibility = 0;
			}
			else
			{
				adornerVisibility = 2;
			}
			AdornerVisibility = (Visibility)adornerVisibility;
		}
	}

	public Visibility AdornerVisibility
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			C(AH.A(68912));
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

	public Arrangement(int intSlides)
	{
		IsChecked = false;
		SlideCount = intSlides;
	}

	private void C(string A)
	{
		PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
		if (propertyChangedEventHandler == null)
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
			propertyChangedEventHandler(this, new PropertyChangedEventArgs(A));
			return;
		}
	}

	internal abstract void A(List<ShapeItem> A, Container B, Preferences C, ref double D);

	internal abstract void B(List<ShapeItem> A, Container B, Preferences C);

	internal void C(Microsoft.Office.Interop.PowerPoint.Shape A, float B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		shape.LockAspectRatio = MsoTriState.msoTrue;
		Microsoft.Office.Interop.PowerPoint.Shape shape2;
		(shape2 = shape).Width = (float)((double)shape2.Width * Math.Sqrt(B / (shape.Width * shape.Height)));
		shape = null;
	}

	internal float C(float A)
	{
		return (float)Math.Round(A, 4);
	}
}
