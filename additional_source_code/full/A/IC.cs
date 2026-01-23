using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace A;

internal sealed class IC<B>
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<UIElement, bool> A;

		public static Func<Rect, double> A;

		public static Func<Rect, double> B;

		public static Func<Rect, double> C;

		public static Func<Rect, double> D;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(UIElement A)
		{
			return A == null;
		}

		[SpecialName]
		internal double A(Rect A)
		{
			return A.Left;
		}

		[SpecialName]
		internal double B(Rect A)
		{
			return A.Top;
		}

		[SpecialName]
		internal double C(Rect A)
		{
			return A.Right;
		}

		[SpecialName]
		internal double D(Rect A)
		{
			return A.Bottom;
		}
	}

	[CompilerGenerated]
	internal sealed class HC
	{
		public ScrollViewer A;

		[SpecialName]
		internal Rect A(UIElement A)
		{
			return IC<B>.A(this.A, A);
		}
	}

	private readonly Func<ListBox> m_A;

	private readonly bool m_A;

	private readonly double? m_A;

	private ScrollViewer m_A;

	private DateTime? m_A;

	private ListBox m_A;

	private bool m_B;

	private DateTime? m_B;

	private const double m_A = 5.0;

	private Visibility? m_A;

	[CompilerGenerated]
	private bool C;

	private DateTime? C;

	private ListBox A
	{
		get
		{
			if (this.A == null)
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
				this.A = this.A();
			}
			return this.A;
		}
	}

	internal bool OverrideInteractedRecently
	{
		[CompilerGenerated]
		get
		{
			return this.C;
		}
		[CompilerGenerated]
		set
		{
			this.C = value;
		}
	}

	internal IC(Func<ListBox> A, double? B = null, bool C = false)
	{
		this.B = true;
		this.A = A;
		this.A = B;
		this.A = C;
	}

	private bool A()
	{
		if (this.A == null)
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
			this.A = GetScrollViewer(this.A);
			if (this.A != null)
			{
				this.A.ScrollChanged += A;
				this.A.PreviewMouseLeftButtonDown += A;
			}
		}
		return this.A != null;
	}

	internal void A()
	{
		if (A())
		{
			this.A.ScrollToTop();
		}
	}

	internal void A(List<B> A)
	{
		if (A == null || A.Count == 0)
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
			if (this.A.HasValue)
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
				DateTime? a = this.A;
				DateTime t = DateTime.UtcNow.AddMilliseconds(0.0 - this.A.Value);
				bool? obj;
				if (!a.HasValue)
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
					obj = null;
				}
				else
				{
					obj = DateTime.Compare(a.GetValueOrDefault(), t) > 0;
				}
				bool? flag = obj;
				if (flag == true)
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
					break;
				}
				this.A = DateTime.UtcNow;
			}
			if (!this.A())
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
				if (!this.A)
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
					if (!OverrideInteractedRecently)
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
						if (this.B())
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									return;
								}
							}
						}
					}
					else
					{
						OverrideInteractedRecently = false;
					}
				}
				List<UIElement> list = A.Select([SpecialName] (B val) => IC<B>.A(this.A, val)).ToList();
				Func<UIElement, bool> predicate;
				if (_Closure_0024__.A == null)
				{
					predicate = (_Closure_0024__.A = [SpecialName] (UIElement uIElement) => uIElement == null);
				}
				else
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
					predicate = _Closure_0024__.A;
				}
				if (!list.Any(predicate))
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
					if (IC<B>.A(this.A, IC<B>.A(this.A, list)) == 0.0)
					{
						return;
					}
				}
				this.B = false;
				try
				{
					using List<B>.Enumerator enumerator = A.GetEnumerator();
					while (enumerator.MoveNext())
					{
						B current = enumerator.Current;
						this.A.UpdateLayout();
						this.A.ScrollIntoView(current);
						JH.A();
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				finally
				{
					this.B = true;
				}
			}
		}
	}

	private static double A(ScrollViewer A, Rect B)
	{
		Rect rect = new Rect(new Point(0.0, 0.0), A.RenderSize);
		if (rect.Top <= B.Top && B.Bottom <= rect.Bottom)
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
					return 0.0;
				}
			}
		}
		bool flag = B.Height < rect.Height;
		double num = B.Bottom - rect.Bottom;
		double num2 = B.Top - rect.Top;
		double result;
		if (!(B.Top >= rect.Top))
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
			if (!flag)
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
				result = num;
			}
			else
			{
				result = num2;
			}
		}
		else if (!flag)
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
			result = num2;
		}
		else
		{
			result = num;
		}
		return result;
	}

	private static Rect A(ScrollViewer A, List<UIElement> B)
	{
		List<Rect> source = B.Select([SpecialName] (UIElement b) => IC<B>.A(A, b)).ToList();
		Point point = new Point(source.Min([SpecialName] (Rect rect) => rect.Left), source.Min([SpecialName] (Rect rect) => rect.Top));
		Func<Rect, double> selector;
		if (_Closure_0024__.C == null)
		{
			selector = (_Closure_0024__.C = [SpecialName] (Rect rect) => rect.Right);
		}
		else
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			selector = _Closure_0024__.C;
		}
		double x = source.Max(selector);
		Func<Rect, double> selector2;
		if (_Closure_0024__.D == null)
		{
			selector2 = (_Closure_0024__.D = [SpecialName] (Rect rect) => rect.Bottom);
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
			selector2 = _Closure_0024__.D;
		}
		Point point2 = new Point(x, source.Max(selector2));
		return new Rect(point, point2);
	}

	private static A A<A>(DependencyObject A) where A : DependencyObject
	{
		checked
		{
			int num = VisualTreeHelper.GetChildrenCount(A) - 1;
			for (int i = 0; i <= num; i++)
			{
				DependencyObject child = VisualTreeHelper.GetChild(A, i);
				if (child is A)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							return (A)child;
						}
					}
				}
				if (child == null)
				{
					continue;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				A val = IC<B>.A<A>(child);
				if (val == null)
				{
					continue;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					return val;
				}
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				return null;
			}
		}
	}

	private static UIElement A(ListBox A, object B)
	{
		return (UIElement)A.ItemContainerGenerator.ContainerFromItem(RuntimeHelpers.GetObjectValue(B));
	}

	private static Rect A(ScrollViewer A, UIElement B)
	{
		return B.TransformToAncestor(A).TransformBounds(new Rect(new Point(0.0, 0.0), B.RenderSize));
	}

	private static ScrollViewer GetScrollViewer(ListBox lbox)
	{
		return A<ScrollViewer>(lbox);
	}

	private void A(object A, ScrollChangedEventArgs B)
	{
		if (this.B)
		{
			this.B = DateTime.UtcNow;
		}
		Visibility computedVerticalScrollBarVisibility = this.A.ComputedVerticalScrollBarVisibility;
		if (object.Equals(this.A, computedVerticalScrollBarVisibility))
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
			ScrollViewer a = this.A;
			Thickness padding;
			if (computedVerticalScrollBarVisibility != Visibility.Visible)
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
				padding = new Thickness(0.0);
			}
			else
			{
				padding = new Thickness(0.0, 0.0, 8.0, 0.0);
			}
			a.Padding = padding;
			this.A = computedVerticalScrollBarVisibility;
			return;
		}
	}

	private void A(object A, MouseButtonEventArgs B)
	{
		C = DateTime.UtcNow;
	}

	private bool B()
	{
		DateTime? dateTime = this.B;
		DateTime? c;
		if (dateTime.HasValue)
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
			DateTime? dateTime2 = dateTime;
			c = C;
			bool? obj;
			if (!(dateTime2.HasValue & c.HasValue))
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
				obj = null;
			}
			else
			{
				obj = DateTime.Compare(dateTime2.GetValueOrDefault(), c.GetValueOrDefault()) < 0;
			}
			bool? flag = obj;
			if (flag != true)
			{
				goto IL_0097;
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
		}
		dateTime = C;
		goto IL_0097;
		IL_0097:
		c = dateTime;
		DateTime t = DateTime.UtcNow.AddSeconds(-5.0);
		return object.Equals(c.HasValue ? new bool?(DateTime.Compare(c.GetValueOrDefault(), t) >= 0) : ((bool?)null), true);
	}

	[SpecialName]
	[CompilerGenerated]
	private UIElement A(B A)
	{
		return IC<B>.A(this.A, A);
	}
}
