using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using A;

namespace PowerPointAddIn1.Explorer;

public sealed class Options
{
	[CompilerGenerated]
	private static EventHandler<PropertyChangedEventArgs> A;

	private static string A = "";

	private static bool A;

	private static bool B;

	private static bool C;

	private static bool D;

	private static bool E;

	private static bool F;

	private static bool G;

	private static bool H;

	private static bool I;

	private static bool J;

	private static bool K;

	private static bool L;

	private static bool M;

	public static string SearchQuery
	{
		get
		{
			return Options.A;
		}
		set
		{
			Options.A = value;
			RaiseStaticPropertyChanged(AH.A(115874));
		}
	}

	public static bool ShowPreviews
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
			RaiseStaticPropertyChanged(AH.A(115897));
		}
	}

	public static bool ShowAll
	{
		get
		{
			return B;
		}
		set
		{
			B = value;
			RaiseStaticPropertyChanged(AH.A(115922));
		}
	}

	public static bool ShowCharts
	{
		get
		{
			return C;
		}
		set
		{
			C = value;
			RaiseStaticPropertyChanged(AH.A(115937));
		}
	}

	public static bool ShowTables
	{
		get
		{
			return D;
		}
		set
		{
			D = value;
			RaiseStaticPropertyChanged(AH.A(115958));
		}
	}

	public static bool ShowImages
	{
		get
		{
			return E;
		}
		set
		{
			E = value;
			RaiseStaticPropertyChanged(AH.A(115979));
		}
	}

	public static bool ShowMedia
	{
		get
		{
			return F;
		}
		set
		{
			F = value;
			RaiseStaticPropertyChanged(AH.A(116000));
		}
	}

	public static bool ShowInk
	{
		get
		{
			return G;
		}
		set
		{
			G = value;
			RaiseStaticPropertyChanged(AH.A(116019));
		}
	}

	public static bool ShowSmartArt
	{
		get
		{
			return H;
		}
		set
		{
			H = value;
			RaiseStaticPropertyChanged(AH.A(116034));
		}
	}

	public static bool ShowEmbeddedExcel
	{
		get
		{
			return I;
		}
		set
		{
			I = value;
			RaiseStaticPropertyChanged(AH.A(116059));
		}
	}

	public static bool ShowEmbeddedWord
	{
		get
		{
			return J;
		}
		set
		{
			J = value;
			RaiseStaticPropertyChanged(AH.A(116094));
		}
	}

	public static bool ShowComments
	{
		get
		{
			return K;
		}
		set
		{
			K = value;
			RaiseStaticPropertyChanged(AH.A(116127));
		}
	}

	public static bool ShowNotes
	{
		get
		{
			return L;
		}
		set
		{
			L = value;
			RaiseStaticPropertyChanged(AH.A(116152));
		}
	}

	public static bool ShowHyperlinks
	{
		get
		{
			return M;
		}
		set
		{
			M = value;
			RaiseStaticPropertyChanged(AH.A(116171));
		}
	}

	public static event EventHandler<PropertyChangedEventArgs> StaticPropertyChanged
	{
		[CompilerGenerated]
		add
		{
			EventHandler<PropertyChangedEventArgs> eventHandler = Options.A;
			EventHandler<PropertyChangedEventArgs> eventHandler2;
			do
			{
				eventHandler2 = eventHandler;
				EventHandler<PropertyChangedEventArgs> value2 = (EventHandler<PropertyChangedEventArgs>)Delegate.Combine(eventHandler2, value);
				eventHandler = Interlocked.CompareExchange(ref Options.A, value2, eventHandler2);
			}
			while ((object)eventHandler != eventHandler2);
		}
		[CompilerGenerated]
		remove
		{
			EventHandler<PropertyChangedEventArgs> eventHandler = Options.A;
			EventHandler<PropertyChangedEventArgs> eventHandler2;
			do
			{
				eventHandler2 = eventHandler;
				EventHandler<PropertyChangedEventArgs> value2 = (EventHandler<PropertyChangedEventArgs>)Delegate.Remove(eventHandler2, value);
				eventHandler = Interlocked.CompareExchange(ref Options.A, value2, eventHandler2);
			}
			while ((object)eventHandler != eventHandler2);
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

	public static void RaiseStaticPropertyChanged(string propName)
	{
		EventHandler<PropertyChangedEventArgs> a = Options.A;
		if (a == null)
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
			a(null, new PropertyChangedEventArgs(propName));
			return;
		}
	}
}
