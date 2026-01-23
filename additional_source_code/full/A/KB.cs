using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using ExcelAddIn1.Audit.Check.UI;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

internal sealed class KB
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<object, string> A;

		public static Func<int?, bool> A;

		public static Func<int?, int> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal string A(object A)
		{
			if (A == null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return null;
					}
				}
			}
			return A.ToString();
		}

		[SpecialName]
		internal bool A(int? A)
		{
			return A.HasValue;
		}

		[SpecialName]
		internal int A(int? A)
		{
			return A.Value;
		}
	}

	[CompilerGenerated]
	internal sealed class IB
	{
		public List<string> A;

		public Func<KeyValuePair<string, List<string>>, bool> A;

		public IB(IB A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal bool A(KeyValuePair<string, List<string>> A)
		{
			return A.Value.SequenceEqual(this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class JB
	{
		public int? A;

		[SpecialName]
		internal int? A(string A)
		{
			if (!int.TryParse(A, out var result))
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return this.A;
					}
				}
			}
			return result;
		}
	}

	[CompilerGenerated]
	private static (string CmdName, List<string> ArgStrs) m_A;

	private static Dictionary<string, List<string>> m_A = new Dictionary<string, List<string>>();

	private static int m_A = 0;

	private const string m_A = "yyyy-MM-dd HH:mm:ss.fff";

	private static readonly string m_B = VH.A(8270);

	private static readonly string m_C = VH.A(8375);

	private static readonly string m_D = VH.A(8396);

	private static readonly object m_A = RuntimeHelpers.GetObjectValue(new object());

	private static char m_A = '|';

	private const char m_B = ',';

	private const char m_C = '|';

	internal static (string CmdName, List<string> ArgStrs) StartupCmdInfo
	{
		[CompilerGenerated]
		get
		{
			return KB.m_A;
		}
		[CompilerGenerated]
		set
		{
			KB.m_A = value;
		}
	}

	internal static string A(string A, List<int> B)
	{
		object a = KB.m_A;
		ObjectFlowControl.CheckForSyncLockOnValueType(a);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(a, ref lockTaken);
			object a2;
			if (B == null)
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
				a2 = null;
			}
			else
			{
				a2 = B.Cast<object>();
			}
			string text = KB.A((IEnumerable<object>)a2);
			return KB.A(KB.m_B, A, text);
		}
		finally
		{
			if (lockTaken)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					Monitor.Exit(a);
					break;
				}
			}
		}
	}

	private static string A(string A, params object[] B)
	{
		int b = KB.A(A);
		KB.A();
		return KB.A(A, b, B);
	}

	private static void A()
	{
		while (!modFunctionsStr.IsBlank(A(A: false).CmdName))
		{
			Thread.Sleep(1000);
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
			return;
		}
	}

	private static (string CmdName, List<string> ArgStrs) A(bool A = false)
	{
		IB a = default(IB);
		IB CS_0024_003C_003E8__locals7 = new IB(a);
		List<string> item = new List<string>();
		CS_0024_003C_003E8__locals7.A = KB.A();
		if (A)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			KB.A("", 0);
			while (true)
			{
				Dictionary<string, List<string>> a2 = KB.m_A;
				Func<KeyValuePair<string, List<string>>, bool> predicate;
				if (CS_0024_003C_003E8__locals7.A != null)
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
					predicate = CS_0024_003C_003E8__locals7.A;
				}
				else
				{
					predicate = (CS_0024_003C_003E8__locals7.A = [SpecialName] (KeyValuePair<string, List<string>> keyValuePair2) => keyValuePair2.Value.SequenceEqual(CS_0024_003C_003E8__locals7.A));
				}
				KeyValuePair<string, List<string>> keyValuePair = a2.FirstOrDefault(predicate);
				if (keyValuePair.Key == null)
				{
					break;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_0091;
					}
					continue;
					end_IL_0091:
					break;
				}
				KB.m_A.Remove(keyValuePair.Key);
			}
		}
		string text = CS_0024_003C_003E8__locals7.A[0];
		(string, List<string>) result;
		if (modFunctionsStr.IsBlank(text))
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
			result = ("", item);
		}
		else
		{
			string text2 = CS_0024_003C_003E8__locals7.A[1];
			DateTime result2;
			if (modFunctionsStr.IsBlank(text2))
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
				result = ("", item);
			}
			else if (!DateTime.TryParseExact(text2, VH.A(8223), CultureInfo.InvariantCulture, DateTimeStyles.None, out result2))
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
				result = ("", item);
			}
			else if (DateTime.Compare(DateTime.UtcNow, result2) > 0)
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
				result = ("", item);
			}
			else
			{
				int num = text.IndexOf(KB.m_A);
				if (num < 1)
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
					result = ("", item);
				}
				else
				{
					string text3 = text.Substring(0, num);
					if (modFunctionsStr.IsBlank(text3))
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
						result = ("", item);
					}
					else
					{
						item = KB.A(text.Substring(checked(num + 1)));
						result = (text3, item);
					}
				}
			}
		}
		return result;
	}

	private static int A(string A)
	{
		if (object.Equals(A, VH.A(8270)))
		{
			return 60;
		}
		return 20;
	}

	private static string A(string A, int B, params object[] C)
	{
		string text = "";
		string text2 = "";
		if (!modFunctionsStr.IsBlank(A))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			text = string.Format(VH.A(8301), A, KB.m_A, KB.A(C));
			text2 = DateTime.UtcNow.AddSeconds(B).ToString(VH.A(8223));
		}
		KB.B(KB.m_C, text);
		KB.B(KB.m_D, text2);
		string text3 = "";
		checked
		{
			if (!modFunctionsStr.IsBlank(text))
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
				KB.m_A++;
				text3 = KB.m_A.ToString();
				KB.m_A[text3] = new List<string> { text, text2 };
			}
			return text3;
		}
	}

	private static string A(IEnumerable<object> A)
	{
		if (A == null)
		{
			return null;
		}
		Func<object, string> selector;
		if (_Closure_0024__.A == null)
		{
			selector = (_Closure_0024__.A = [SpecialName] (object obj) =>
			{
				if (obj == null)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							return (string)null;
						}
					}
				}
				return obj.ToString();
			});
		}
		else
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			selector = _Closure_0024__.A;
		}
		return modFunctionsStr.JoinEscString(A.Select(selector), ',', '|');
	}

	private static List<string> A(string A)
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
					return null;
				}
			}
		}
		return modFunctionsStr.SplitEscString(A, ',', '|');
	}

	private static List<int> A(string A)
	{
		int? A2 = null;
		List<string> list = KB.A(A);
		if (list == null)
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
					return null;
				}
			}
		}
		IEnumerable<int?> source = from num in list.Select([SpecialName] (string s) =>
			{
				if (!int.TryParse(s, out var result))
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							return A2;
						}
					}
				}
				return result;
			})
			where num.HasValue
			select num;
		Func<int?, int> selector;
		if (_Closure_0024__.A == null)
		{
			selector = (_Closure_0024__.A = [SpecialName] (int? num) => num.Value);
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
			selector = _Closure_0024__.A;
		}
		return source.Select(selector).ToList();
	}

	private static void B(string A, string B)
	{
		if (modFunctionsStr.IsBlank(B))
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
					clsRegistry.DeleteRegistryKey(A);
					return;
				}
			}
		}
		clsRegistry.SetRegistryValue(A, B);
	}

	internal static void A(string A)
	{
		if (modFunctionsStr.IsBlank(A))
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
			object a = KB.m_A;
			ObjectFlowControl.CheckForSyncLockOnValueType(a);
			bool lockTaken = false;
			try
			{
				Monitor.Enter(a, ref lockTaken);
				List<string> value = new List<string>();
				if (!KB.m_A.TryGetValue(A, out value))
				{
					return;
				}
				if (!value.SequenceEqual(KB.A()))
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				KB.A("", 0);
				KB.m_A.Remove(A);
				return;
			}
			finally
			{
				if (lockTaken)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						Monitor.Exit(a);
						break;
					}
				}
			}
		}
	}

	private static List<string> A()
	{
		string registryValue = clsRegistry.GetRegistryValue(KB.m_C);
		string registryValue2 = clsRegistry.GetRegistryValue(KB.m_D);
		return new List<string> { registryValue, registryValue2 };
	}

	internal static void C()
	{
		object a = KB.m_A;
		ObjectFlowControl.CheckForSyncLockOnValueType(a);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(a, ref lockTaken);
			StartupCmdInfo = A(A: true);
			if (Operators.CompareString(StartupCmdInfo.CmdName, KB.m_B, TextCompare: false) != 0)
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
				List<string> item = StartupCmdInfo.ArgStrs;
				if (item.Count < 2)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				string a2 = item[0];
				string b = item[1];
				D(a2, b);
				return;
			}
		}
		finally
		{
			if (lockTaken)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					Monitor.Exit(a);
					break;
				}
			}
		}
	}

	private static void D(string A, string B)
	{
		try
		{
			MH.A.Application.Workbooks.Open(A, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(string.Format(VH.A(8320), modFunctionsException.DetailedExcMessage(false, new Exception[1] { ex2 })));
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
			return;
		}
		finally
		{
		}
		wpfAudit.RunCheckOnOpen = (YesNo: true, SheetIndexesList: KB.A(B));
		wpfAudit.AssumeAlreadySavedOnNextRun = true;
		Pane.Toggle(blnPressed: true);
	}

	internal static bool A(string A)
	{
		object a = KB.m_A;
		ObjectFlowControl.CheckForSyncLockOnValueType(a);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(a, ref lockTaken);
			int result;
			if (object.Equals(StartupCmdInfo.CmdName, KB.m_B))
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				result = ((string.Compare(StartupCmdInfo.ArgStrs.FirstOrDefault(), A, ignoreCase: true) == 0) ? 1 : 0);
			}
			else
			{
				result = 0;
			}
			return (byte)result != 0;
		}
		finally
		{
			if (lockTaken)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					Monitor.Exit(a);
					break;
				}
			}
		}
	}
}
