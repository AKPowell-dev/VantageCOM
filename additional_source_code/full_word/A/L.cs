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

[GeneratedCode("MyTemplate", "11.0.0.0")]
[HideModuleName]
[StandardModule]
internal sealed class L
{
	[EditorBrowsable(EditorBrowsableState.Never)]
	[MyGroupCollection("System.Web.Services.Protocols.SoapHttpClientProtocol", "Create__Instance__", "Dispose__Instance__", "")]
	internal sealed class J
	{
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DebuggerHidden]
		public J()
		{
		}

		[DebuggerHidden]
		[EditorBrowsable(EditorBrowsableState.Never)]
		public override bool Equals(object o)
		{
			return base.Equals(RuntimeHelpers.GetObjectValue(o));
		}

		[DebuggerHidden]
		[EditorBrowsable(EditorBrowsableState.Never)]
		public override int GetHashCode()
		{
			return base.GetHashCode();
		}

		[DebuggerHidden]
		[EditorBrowsable(EditorBrowsableState.Never)]
		internal Type A()
		{
			return typeof(J);
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
					switch (6)
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
	internal sealed class K<A> where A : new()
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
					val = new A();
					this.A.Value = val;
				}
				return val;
			}
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		[DebuggerHidden]
		public K()
		{
			this.A = new ContextValue<A>();
		}
	}

	private static readonly K<I> m_A = new K<I>();

	private static readonly K<H> m_A = new K<H>();

	private static readonly K<User> m_A = new K<User>();

	private static readonly K<J> m_A = new K<J>();

	[HelpKeyword("My.Computer")]
	internal static I A
	{
		[DebuggerHidden]
		get
		{
			return L.m_A.A;
		}
	}

	[HelpKeyword("My.Application")]
	internal static H A
	{
		[DebuggerHidden]
		get
		{
			return L.m_A.A;
		}
	}

	[HelpKeyword("My.User")]
	internal static User A
	{
		[DebuggerHidden]
		get
		{
			return L.m_A.A;
		}
	}

	[HelpKeyword("My.WebServices")]
	internal static J A
	{
		[DebuggerHidden]
		get
		{
			return L.m_A.A;
		}
	}
}
