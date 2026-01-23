using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Media;
using A;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Template.Wizard;

public sealed class TemplateWizardLayout : INotifyPropertyChanged
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private CustomLayout m_A;

	private string m_A;

	private int m_A;

	private Color m_A;

	private Color B;

	public CustomLayout Layout
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(121018));
		}
	}

	public string Name
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(63335));
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
			if (value == 0)
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
				object obj = ColorConverter.ConvertFromString(AH.A(121031));
				FontColor = ((obj != null) ? ((Color)obj) : default(Color));
				FillColor = System.Windows.Media.Colors.White;
			}
			else
			{
				object obj2 = ColorConverter.ConvertFromString(AH.A(121046));
				FontColor = ((obj2 != null) ? ((Color)obj2) : default(Color));
				FillColor = System.Windows.Media.Colors.White;
			}
			A(AH.A(116304));
		}
	}

	public Color FontColor
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(113606));
		}
	}

	public Color FillColor
	{
		get
		{
			return B;
		}
		set
		{
			B = value;
			A(AH.A(49852));
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
	}

	public TemplateWizardLayout(CustomLayout lay)
	{
		Layout = lay;
		Name = lay.Name;
		switch (Helpers.GetLayoutType(lay))
		{
		case SlideType.Title:
			Index = 1;
			break;
		case SlideType.TableOfContents:
		case SlideType.Agenda:
			Index = 2;
			break;
		case SlideType.Flysheet:
			Index = 3;
			break;
		case SlideType.Legal:
			Index = 4;
			break;
		case SlideType.Contact:
			Index = 5;
			break;
		case SlideType.Blank:
			Index = 6;
			break;
		case SlideType.CoverFront:
			Index = 7;
			break;
		case SlideType.CoverBack:
			Index = 8;
			break;
		default:
			Index = 0;
			break;
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
			switch (4)
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
