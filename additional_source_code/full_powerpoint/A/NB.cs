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

[StandardModule]
[HideModuleName]
[GeneratedCode("MyTemplate", "11.0.0.0")]
internal sealed class NB
{
	[EditorBrowsable(EditorBrowsableState.Never)]
	[MyGroupCollection("System.Web.Services.Protocols.SoapHttpClientProtocol", "Create__Instance__", "Dispose__Instance__", "")]
	internal sealed class LB
	{
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DebuggerHidden]
		public LB()
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
			return typeof(LB);
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
				return new A();
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
	internal sealed class MB<A> where A : new()
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
					val = new A();
					this.A.Value = val;
				}
				return val;
			}
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		[DebuggerHidden]
		public MB()
		{
			this.A = new ContextValue<A>();
		}
	}

	private static readonly MB<KB> m_A = new MB<KB>();

	private static readonly MB<JB> m_A = new MB<JB>();

	private static readonly MB<User> m_A = new MB<User>();

	private static readonly MB<LB> m_A = new MB<LB>();

	[HelpKeyword("My.Computer")]
	internal static KB A
	{
		[DebuggerHidden]
		get
		{
			return NB.m_A.A;
		}
	}

	[HelpKeyword("My.Application")]
	internal static JB A
	{
		[DebuggerHidden]
		get
		{
			return NB.m_A.A;
		}
	}

	[HelpKeyword("My.User")]
	internal static User A
	{
		[DebuggerHidden]
		get
		{
			return NB.m_A.A;
		}
	}

	[HelpKeyword("My.WebServices")]
	internal static LB A
	{
		[DebuggerHidden]
		get
		{
			return NB.m_A.A;
		}
	}
}
