using System;
using System.CodeDom.Compiler;
using System.Diagnostics;
using Macabacus_Word;
using Microsoft.Office.Tools.Word;

namespace A;

[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
[DebuggerNonUserCode]
internal sealed class PC
{
	private static ThisAddIn m_A;

	private static ApplicationFactory m_A;

	private static QC m_A;

	internal static ThisAddIn A
	{
		get
		{
			return PC.m_A;
		}
		set
		{
			if (PC.m_A == null)
			{
				PC.m_A = value;
				return;
			}
			throw new NotSupportedException();
		}
	}

	internal static ApplicationFactory A
	{
		get
		{
			return PC.m_A;
		}
		set
		{
			if (PC.m_A == null)
			{
				PC.m_A = value;
				return;
			}
			throw new NotSupportedException();
		}
	}

	internal static QC A
	{
		get
		{
			if (PC.m_A == null)
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
				PC.m_A = new QC(PC.m_A.GetRibbonFactory());
			}
			return PC.m_A;
		}
	}

	private PC()
	{
	}
}
