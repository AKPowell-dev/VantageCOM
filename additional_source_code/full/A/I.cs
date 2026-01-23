using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.ApplicationServices;
using Microsoft.VisualBasic.CompilerServices;
using Microsoft.VisualBasic.MyServices.Internal;

namespace A;

[HideModuleName]
[GeneratedCode("MyTemplate", "11.0.0.0")]
[StandardModule]
internal sealed class I
{
	[MyGroupCollection("System.Web.Services.Protocols.SoapHttpClientProtocol", "Create__Instance__", "Dispose__Instance__", "")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	internal sealed class G
	{
		[DebuggerHidden]
		[EditorBrowsable(EditorBrowsableState.Never)]
		public G()
		{
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		[DebuggerHidden]
		public override bool Equals(object o)
		{
			return base.Equals(RuntimeHelpers.GetObjectValue(o));
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		[DebuggerHidden]
		public override int GetHashCode()
		{
			return base.GetHashCode();
		}

		[DebuggerHidden]
		[EditorBrowsable(EditorBrowsableState.Never)]
		internal Type A()
		{
			return typeof(G);
		}

		[DebuggerHidden]
		[EditorBrowsable(EditorBrowsableState.Never)]
		public override string ToString()
		{
			return base.ToString();
		}

		[DebuggerHidden]
		private static A A<A>(A A) where A : new()
		{
			if (A == null)
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
						return new A();
					}
				}
			}
			return A;
		}

		[DebuggerHidden]
		private void A<A>(ref A A)
		{
			A = default(A);
		}
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[ComVisible(false)]
	internal sealed class H<A> where A : new()
	{
		private readonly ContextValue<A> m_A;

		internal A A
		{
			[DebuggerHidden]
			get
			{
				A val = this.A.Value;
				if (val == null)
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
					val = new A();
					this.A.Value = val;
				}
				return val;
			}
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		[DebuggerHidden]
		public H()
		{
			this.A = new ContextValue<A>();
		}
	}

	private static readonly H<F> m_A = new H<F>();

	private static readonly H<E> m_A = new H<E>();

	private static readonly H<User> m_A = new H<User>();

	private static readonly H<G> m_A = new H<G>();

	[HelpKeyword("My.Computer")]
	internal static F A
	{
		[DebuggerHidden]
		get
		{
			return I.m_A.A;
		}
	}

	[HelpKeyword("My.Application")]
	internal static E A
	{
		[DebuggerHidden]
		get
		{
			return I.m_A.A;
		}
	}

	[HelpKeyword("My.User")]
	internal static User A
	{
		[DebuggerHidden]
		get
		{
			return I.m_A.A;
		}
	}

	[HelpKeyword("My.WebServices")]
	internal static G A
	{
		[DebuggerHidden]
		get
		{
			return I.m_A.A;
		}
	}
}
