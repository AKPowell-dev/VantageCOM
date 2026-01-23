using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using ExcelAddIn1.Format;
using ExcelAddIn1.Keyboard;
using ExcelAddIn1.Library2.Versioning;
using ExcelAddIn1.RowsColumns;
using ExcelAddIn1.View;
using MacabacusMacros;
using MacabacusMacros.Config.Settings;
using MacabacusMacros.UI;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

public sealed class clsSettings
{
	[CompilerGenerated]
	private XmlElement m_A;

	[CompilerGenerated]
	private bool m_A;

	[CompilerGenerated]
	private bool B;

	[CompilerGenerated]
	private bool C;

	private List<string> m_A;

	private ColorCycle m_A;

	private ColorCycle B;

	private List<string> B;

	private ColorCycle C;

	private ColorCycle D;

	private List<string> C;

	private List<string> D;

	private List<string> E;

	[CompilerGenerated]
	private int m_A;

	[CompilerGenerated]
	private int B;

	[CompilerGenerated]
	private int C;

	private List<string> F;

	private NumberFormatCycle m_A;

	private NumberFormatCycle B;

	private NumberFormatCycle C;

	private NumberFormatCycle D;

	private NumberFormatCycle E;

	private NumberFormatCycle F;

	private NumberFormatCycle G;

	private List<string> G;

	private List<float> m_A;

	[CompilerGenerated]
	private int D;

	[CompilerGenerated]
	private int E;

	[CompilerGenerated]
	private int F;

	private List<float> B;

	private List<float> C;

	[CompilerGenerated]
	private bool D;

	[CompilerGenerated]
	private bool E;

	[CompilerGenerated]
	private string m_A;

	[CompilerGenerated]
	private int G;

	[CompilerGenerated]
	private bool F;

	[CompilerGenerated]
	private bool G;

	[CompilerGenerated]
	private int H;

	[CompilerGenerated]
	private bool H;

	[CompilerGenerated]
	private bool I;

	[CompilerGenerated]
	private bool J;

	[CompilerGenerated]
	private bool K;

	private Dictionary<string, List<XmlNode>> m_A;

	[CompilerGenerated]
	private int I;

	[CompilerGenerated]
	private bool L;

	[CompilerGenerated]
	private bool M;

	[CompilerGenerated]
	private int J;

	public XmlDocument SettingsXml => Manage.GetSettings();

	public XmlElement SettingsRoot
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public bool AutoColorOnEntry
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public bool AutoColorText
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	public bool AutoColorDates
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

	public List<string> AutoColors
	{
		get
		{
			if (this.m_A == null)
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
				this.m_A = new List<string>();
				XmlNode xmlNode = SettingsRoot.SelectSingleNode(VH.A(194853));
				this.m_A.Add(xmlNode.SelectSingleNode(VH.A(194874)).InnerText);
				this.m_A.Add(xmlNode.SelectSingleNode(VH.A(194897)).InnerText);
				this.m_A.Add(xmlNode.SelectSingleNode(VH.A(194922)).InnerText);
				this.m_A.Add(xmlNode.SelectSingleNode(VH.A(194949)).InnerText);
				this.m_A.Add(xmlNode.SelectSingleNode(VH.A(194980)).InnerText);
				this.m_A.Add(xmlNode.SelectSingleNode(VH.A(195009)).InnerText);
				this.m_A.Add(xmlNode.SelectSingleNode(VH.A(195040)).InnerText);
				xmlNode = null;
			}
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public ColorCycle FontColorCycle
	{
		get
		{
			if (this.m_A == null)
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
				this.m_A = new ColorCycle(SettingsXml, VH.A(195071), VH.A(195100));
			}
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public ColorCycle FillColorCycle
	{
		get
		{
			if (this.B == null)
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
				this.B = new ColorCycle(SettingsXml, VH.A(195133), VH.A(195162));
			}
			return this.B;
		}
		set
		{
			this.B = value;
		}
	}

	public List<string> NoAutoColorCycle
	{
		get
		{
			if (this.B == null)
			{
				this.B = new List<string>();
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = SettingsRoot.SelectNodes(VH.A(195195)).GetEnumerator();
					while (enumerator.MoveNext())
					{
						XmlNode xmlNode = (XmlNode)enumerator.Current;
						this.B.Add(xmlNode.InnerText);
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
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
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			return this.B;
		}
		set
		{
			this.B = value;
		}
	}

	public ColorCycle BorderColorCycle
	{
		get
		{
			if (this.C == null)
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
				this.C = new ColorCycle(SettingsXml, VH.A(195264), VH.A(195297));
			}
			return this.C;
		}
		set
		{
			this.C = value;
		}
	}

	public ColorCycle ChartColorCycle
	{
		get
		{
			if (this.D == null)
			{
				this.D = new ColorCycle(SettingsXml, VH.A(195334), VH.A(195365));
			}
			return this.D;
		}
		set
		{
			this.D = value;
		}
	}

	public List<string> ChartSeriesColors
	{
		get
		{
			if (this.C == null)
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
				this.C = new List<string>();
				XmlNodeList xmlNodeList = SettingsRoot.SelectNodes(VH.A(195400));
				foreach (XmlNode item in xmlNodeList)
				{
					this.C.Add(item.InnerText);
				}
				xmlNodeList = null;
			}
			return this.C;
		}
		set
		{
			this.C = value;
		}
	}

	public List<string> RecolorColors
	{
		get
		{
			if (this.D == null)
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
				this.D = new List<string>();
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = SettingsRoot.SelectNodes(VH.A(195447)).GetEnumerator();
					while (enumerator.MoveNext())
					{
						XmlNode xmlNode = (XmlNode)enumerator.Current;
						this.D.Add(xmlNode.InnerText);
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			return this.D;
		}
		set
		{
			this.D = value;
		}
	}

	public List<string> BorderStyleCycle
	{
		get
		{
			if (this.E == null)
			{
				this.E = new List<string>();
				{
					IEnumerator enumerator = SettingsRoot.SelectNodes(VH.A(195486)).GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							XmlNode xmlNode = (XmlNode)enumerator.Current;
							this.E.Add(xmlNode.InnerText);
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
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
			}
			return this.E;
		}
		set
		{
			this.E = value;
		}
	}

	public int DefaultFontColor
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public int DefaultFillColor
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	public int DefaultBorderColor
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

	public List<string> DataFunctions
	{
		get
		{
			if (this.F == null)
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
				this.F = new List<string>();
				this.F.AddRange(Strings.Split(SettingsXml.DocumentElement.SelectSingleNode(VH.A(195531)).InnerText, VH.A(2378)));
			}
			return this.F;
		}
		set
		{
			this.F = value;
		}
	}

	public NumberFormatCycle CycleNumber
	{
		get
		{
			if (this.m_A == null)
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
				this.m_A = new NumberFormatCycle(SettingsXml, VH.A(195580), VH.A(195617));
			}
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public NumberFormatCycle CycleCurrency
	{
		get
		{
			if (this.B == null)
			{
				this.B = new NumberFormatCycle(SettingsXml, VH.A(195658), VH.A(195695));
			}
			return this.B;
		}
		set
		{
			this.B = value;
		}
	}

	public NumberFormatCycle CyclePercent
	{
		get
		{
			if (this.C == null)
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
				this.C = new NumberFormatCycle(SettingsXml, VH.A(195724), VH.A(195749));
			}
			return this.C;
		}
		set
		{
			this.C = value;
		}
	}

	public NumberFormatCycle CycleDate
	{
		get
		{
			if (this.D == null)
			{
				this.D = new NumberFormatCycle(SettingsXml, VH.A(195776), VH.A(195795));
			}
			return this.D;
		}
		set
		{
			this.D = value;
		}
	}

	public NumberFormatCycle CycleMultiple
	{
		get
		{
			if (this.E == null)
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
				this.E = new NumberFormatCycle(SettingsXml, VH.A(195816), VH.A(195843));
			}
			return this.E;
		}
		set
		{
			this.E = value;
		}
	}

	public NumberFormatCycle CycleBinary
	{
		get
		{
			if (this.F == null)
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
				this.F = new NumberFormatCycle(SettingsXml, VH.A(195872), VH.A(195895));
			}
			return this.F;
		}
		set
		{
			this.F = value;
		}
	}

	public NumberFormatCycle CycleRatio
	{
		get
		{
			if (this.G == null)
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
				this.G = new NumberFormatCycle(SettingsXml, VH.A(195920), VH.A(195941));
			}
			return this.G;
		}
		set
		{
			this.G = value;
		}
	}

	public List<string> FontStyleCycle
	{
		get
		{
			if (this.G == null)
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
				this.G = new List<string>();
				foreach (XmlNode item in SettingsRoot.SelectNodes(VH.A(195964)))
				{
					this.G.Add(item.InnerText);
				}
			}
			return this.G;
		}
		set
		{
			this.G = value;
		}
	}

	public List<float> FontSizeCycle
	{
		get
		{
			if (this.m_A == null)
			{
				this.m_A = new List<float>();
				foreach (XmlNode item in SettingsRoot.SelectNodes(VH.A(196003)))
				{
					this.m_A.Add(float.Parse(item.InnerText, CultureInfo.InvariantCulture));
				}
			}
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public int DependentsTimeout
	{
		[CompilerGenerated]
		get
		{
			return this.D;
		}
		[CompilerGenerated]
		set
		{
			this.D = value;
		}
	}

	public int IndentMaxLeft
	{
		[CompilerGenerated]
		get
		{
			return this.E;
		}
		[CompilerGenerated]
		set
		{
			this.E = value;
		}
	}

	public int IndentMaxRight
	{
		[CompilerGenerated]
		get
		{
			return this.F;
		}
		[CompilerGenerated]
		set
		{
			this.F = value;
		}
	}

	public List<float> ColumnWidthCycle
	{
		get
		{
			if (B == null)
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
				B = new List<float>();
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = SettingsRoot.SelectNodes(VH.A(196040)).GetEnumerator();
					while (enumerator.MoveNext())
					{
						XmlNode xmlNode = (XmlNode)enumerator.Current;
						B.Add(float.Parse(xmlNode.InnerText, CultureInfo.InvariantCulture));
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0081;
						}
						continue;
						end_IL_0081:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			return B;
		}
		set
		{
			B = value;
		}
	}

	public List<float> RowHeightCycle
	{
		get
		{
			if (C == null)
			{
				C = new List<float>();
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = SettingsRoot.SelectNodes(VH.A(196085)).GetEnumerator();
					while (enumerator.MoveNext())
					{
						XmlNode xmlNode = (XmlNode)enumerator.Current;
						C.Add(float.Parse(xmlNode.InnerText, CultureInfo.InvariantCulture));
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
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			return C;
		}
		set
		{
			C = value;
		}
	}

	public bool UniformRangeEditMode
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[CompilerGenerated]
		set
		{
			D = value;
		}
	}

	public bool ErrorValuePrompt
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[CompilerGenerated]
		set
		{
			E = value;
		}
	}

	public string DefaultErrorValue
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public int ZoomInterval
	{
		[CompilerGenerated]
		get
		{
			return this.G;
		}
		[CompilerGenerated]
		set
		{
			this.G = value;
		}
	}

	public bool CommentAuthor
	{
		[CompilerGenerated]
		get
		{
			return F;
		}
		[CompilerGenerated]
		set
		{
			F = value;
		}
	}

	public bool UndoEnabled
	{
		[CompilerGenerated]
		get
		{
			return G;
		}
		[CompilerGenerated]
		set
		{
			G = value;
		}
	}

	public int UndoMaxCells
	{
		[CompilerGenerated]
		get
		{
			return this.H;
		}
		[CompilerGenerated]
		set
		{
			this.H = value;
		}
	}

	public bool UndoBorders
	{
		[CompilerGenerated]
		get
		{
			return H;
		}
		[CompilerGenerated]
		set
		{
			H = value;
		}
	}

	public bool UndoFont
	{
		[CompilerGenerated]
		get
		{
			return this.I;
		}
		[CompilerGenerated]
		set
		{
			this.I = value;
		}
	}

	public bool UndoFill
	{
		[CompilerGenerated]
		get
		{
			return this.J;
		}
		[CompilerGenerated]
		set
		{
			this.J = value;
		}
	}

	public bool UndoAlignment
	{
		[CompilerGenerated]
		get
		{
			return K;
		}
		[CompilerGenerated]
		set
		{
			K = value;
		}
	}

	public Dictionary<string, List<XmlNode>> CustomCycles
	{
		get
		{
			if (this.m_A == null)
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
				this.m_A = new Dictionary<string, List<XmlNode>>();
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = SettingsXml.DocumentElement.SelectNodes(VH.A(161606)).GetEnumerator();
					IEnumerator enumerator2 = default(IEnumerator);
					while (enumerator.MoveNext())
					{
						XmlNode xmlNode = (XmlNode)enumerator.Current;
						List<XmlNode> list = new List<XmlNode>();
						try
						{
							enumerator2 = xmlNode.SelectNodes(VH.A(187074)).GetEnumerator();
							while (enumerator2.MoveNext())
							{
								XmlNode item = (XmlNode)enumerator2.Current;
								list.Add(item);
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_00a7;
								}
								continue;
								end_IL_00a7:
								break;
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									(enumerator2 as IDisposable).Dispose();
									break;
								}
							}
						}
						this.m_A.Add(xmlNode.Attributes[VH.A(67336)].Value, list);
						list = null;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public int PaintbrushesLimit
	{
		[CompilerGenerated]
		get
		{
			return I;
		}
		[CompilerGenerated]
		set
		{
			I = value;
		}
	}

	public bool AutoAlignRightNumbers
	{
		[CompilerGenerated]
		get
		{
			return L;
		}
		[CompilerGenerated]
		set
		{
			L = value;
		}
	}

	public bool AutoItalicizePercentages
	{
		[CompilerGenerated]
		get
		{
			return M;
		}
		[CompilerGenerated]
		set
		{
			M = value;
		}
	}

	public int TranslatorDelay
	{
		[CompilerGenerated]
		get
		{
			return J;
		}
		[CompilerGenerated]
		set
		{
			J = value;
		}
	}

	public clsSettings(XmlDocument xmlSettings)
	{
		this.m_A = null;
		this.m_A = null;
		this.B = null;
		this.B = null;
		this.C = null;
		this.D = null;
		this.C = null;
		this.D = null;
		this.E = null;
		this.F = null;
		this.m_A = null;
		this.B = null;
		this.C = null;
		this.D = null;
		this.E = null;
		this.F = null;
		this.G = null;
		this.G = null;
		this.m_A = null;
		B = null;
		C = null;
		this.m_A = null;
		A(xmlSettings);
	}

	public clsSettings()
	{
		this.m_A = null;
		this.m_A = null;
		this.B = null;
		this.B = null;
		this.C = null;
		this.D = null;
		this.C = null;
		this.D = null;
		this.E = null;
		this.F = null;
		this.m_A = null;
		this.B = null;
		this.C = null;
		this.D = null;
		this.E = null;
		this.F = null;
		this.G = null;
		this.G = null;
		this.m_A = null;
		B = null;
		C = null;
		this.m_A = null;
		A(Manage.GetXml(false));
	}

	private void A(XmlDocument A)
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		XmlDocument xmlDocument = default(XmlDocument);
		XmlElement xmlElement = default(XmlElement);
		XmlNode xmlNode = default(XmlNode);
		XmlNode xmlNode2 = default(XmlNode);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				int num4;
				double? obj;
				double? num5;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					KH.A = null;
					goto IL_0008;
				case 1962:
					{
						num = num2;
						switch (num3)
						{
						case 2:
							break;
						case 1:
							goto IL_06ca;
						default:
							goto end_IL_0000;
						}
						break;
					}
					IL_06ca:
					num4 = num + 1;
					num = 0;
					switch (num4)
					{
					case 1:
						break;
					case 2:
						goto IL_0008;
					case 3:
						goto IL_000f;
					case 4:
						goto IL_0013;
					case 5:
						goto IL_0023;
					case 6:
						goto IL_002d;
					case 7:
						goto IL_0034;
					case 8:
						goto IL_0066;
					case 10:
						goto IL_007e;
					case 11:
						goto IL_00a7;
					case 12:
						goto IL_00ce;
					case 13:
						goto IL_00f5;
					case 14:
						goto IL_0120;
					case 15:
						goto IL_014a;
					case 16:
						goto IL_0174;
					case 17:
						goto IL_0195;
					case 18:
						goto IL_01c1;
					case 19:
						goto IL_01ed;
					case 20:
						goto IL_0225;
					case 21:
						goto IL_0257;
					case 22:
						goto IL_028f;
					case 23:
						goto IL_02b9;
					case 24:
						goto IL_02e5;
					case 25:
						goto IL_030f;
					case 26:
						goto IL_0337;
					case 27:
						goto IL_0363;
					case 28:
						goto IL_038b;
					case 29:
						goto IL_03b3;
					case 30:
						goto IL_03dd;
					case 31:
						goto IL_0408;
					case 32:
						goto IL_0420;
					case 33:
						goto IL_0451;
					case 34:
						goto IL_0480;
					case 35:
						goto IL_04c4;
					case 36:
						goto IL_0506;
					case 37:
						goto IL_054a;
					case 38:
						goto IL_0590;
					case 39:
						goto IL_0593;
					case 40:
						goto IL_05ad;
					case 41:
						goto IL_05d3;
					case 42:
						goto IL_05d6;
					case 43:
						goto IL_064e;
					case 44:
						goto IL_0672;
					case 46:
						goto IL_067d;
					case 45:
					case 47:
						goto IL_0686;
					case 48:
						goto IL_06aa;
					case 49:
						goto IL_06ad;
					case 51:
						goto end_IL_0000_2;
					default:
						goto end_IL_0000;
					case 9:
					case 50:
					case 52:
						goto end_IL_0000_3;
					}
					goto default;
					IL_0008:
					ProjectData.ClearProjectError();
					num3 = 2;
					goto IL_000f;
					IL_000f:
					num2 = 3;
					xmlDocument = A;
					goto IL_0013;
					IL_0013:
					num2 = 4;
					SettingsRoot = xmlDocument.DocumentElement;
					goto IL_0023;
					IL_0023:
					num2 = 5;
					xmlElement = xmlDocument.DocumentElement;
					goto IL_002d;
					IL_002d:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0034;
					IL_0034:
					num2 = 7;
					if (Operators.CompareString(xmlElement.Name, VH.A(196128), TextCompare: false) != 0)
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
						goto IL_0066;
					}
					goto IL_007e;
					IL_067d:
					num2 = 46;
					Gridlines.AutoHide(blnEnabled: false);
					goto IL_0686;
					IL_064e:
					num2 = 43;
					if (Conversions.ToBoolean(xmlElement.SelectSingleNode(VH.A(141667)).InnerText))
					{
						goto IL_0672;
					}
					goto IL_067d;
					IL_0672:
					num2 = 44;
					Gridlines.AutoHide(blnEnabled: true);
					goto IL_0686;
					IL_0066:
					num2 = 8;
					Forms.ErrorMessage(VH.A(196163));
					goto end_IL_0000_3;
					IL_007e:
					num2 = 10;
					DisabledKeys.DisableKeyF1 = Conversions.ToBoolean(xmlElement.SelectSingleNode(VH.A(161174)).InnerText);
					goto IL_00a7;
					IL_00a7:
					num2 = 11;
					DisabledKeys.DisableKeyInsert = Conversions.ToBoolean(xmlElement.SelectSingleNode(VH.A(161193)).InnerText);
					goto IL_00ce;
					IL_00ce:
					num2 = 12;
					DisabledKeys.DisableKeyNumLock = Conversions.ToBoolean(xmlElement.SelectSingleNode(VH.A(161220)).InnerText);
					goto IL_00f5;
					IL_00f5:
					num2 = 13;
					DisabledKeys.DisableKeyScrollLock = Conversions.ToBoolean(xmlElement.SelectSingleNode(VH.A(161249)).InnerText);
					goto IL_0120;
					IL_0120:
					num2 = 14;
					ZoomInterval = Conversions.ToInteger(xmlElement.SelectSingleNode(VH.A(196268)).InnerText);
					goto IL_014a;
					IL_014a:
					num2 = 15;
					ErrorValuePrompt = Conversions.ToBoolean(xmlElement.SelectSingleNode(VH.A(196293)).InnerText);
					goto IL_0174;
					IL_0174:
					num2 = 16;
					DefaultErrorValue = xmlElement.SelectSingleNode(VH.A(196326)).InnerText;
					goto IL_0195;
					IL_0195:
					num2 = 17;
					CommentAuthor = Conversions.ToBoolean(xmlElement.SelectSingleNode(VH.A(196361)).InnerText);
					goto IL_01c1;
					IL_01c1:
					num2 = 18;
					UniformRangeEditMode = Conversions.ToBoolean(xmlElement.SelectSingleNode(VH.A(196392)).InnerText);
					goto IL_01ed;
					IL_01ed:
					num2 = 19;
					DefaultFontColor = clsColors.RGB2Ole(xmlElement.SelectSingleNode(Constants.XML_DEFAULT_COLORS + VH.A(196433)).InnerText);
					goto IL_0225;
					IL_0225:
					num2 = 20;
					DefaultFillColor = clsColors.RGB2Ole(xmlElement.SelectSingleNode(Constants.XML_DEFAULT_COLORS + VH.A(196454)).InnerText);
					goto IL_0257;
					IL_0257:
					num2 = 21;
					DefaultBorderColor = clsColors.RGB2Ole(xmlElement.SelectSingleNode(Constants.XML_DEFAULT_COLORS + VH.A(196475)).InnerText);
					goto IL_028f;
					IL_028f:
					num2 = 22;
					AutoColorOnEntry = Conversions.ToBoolean(xmlElement.SelectSingleNode(VH.A(186672)).InnerText);
					goto IL_02b9;
					IL_02b9:
					num2 = 23;
					AutoColorText = Conversions.ToBoolean(xmlElement.SelectSingleNode(VH.A(196500)).InnerText);
					goto IL_02e5;
					IL_02e5:
					num2 = 24;
					AutoColorDates = Conversions.ToBoolean(xmlElement.SelectSingleNode(VH.A(196527)).InnerText);
					goto IL_030f;
					IL_030f:
					num2 = 25;
					AutoAlignRightNumbers = Conversions.ToBoolean(xmlElement.SelectSingleNode(VH.A(196556)).InnerText);
					goto IL_0337;
					IL_0337:
					num2 = 26;
					AutoItalicizePercentages = Conversions.ToBoolean(xmlElement.SelectSingleNode(VH.A(196599)).InnerText);
					goto IL_0363;
					IL_0363:
					num2 = 27;
					PaintbrushesLimit = Conversions.ToInteger(xmlElement.SelectSingleNode(VH.A(196648)).InnerText);
					goto IL_038b;
					IL_038b:
					num2 = 28;
					IndentMaxLeft = Conversions.ToInteger(xmlElement.SelectSingleNode(VH.A(196679)).InnerText);
					goto IL_03b3;
					IL_03b3:
					num2 = 29;
					IndentMaxRight = Conversions.ToInteger(xmlElement.SelectSingleNode(VH.A(196706)).InnerText);
					goto IL_03dd;
					IL_03dd:
					num2 = 30;
					AutoFit.Behavior = (AutoFit.HiddenBehavior)Conversions.ToInteger(xmlElement.SelectSingleNode(VH.A(196735)).InnerText);
					goto IL_0408;
					IL_0408:
					num2 = 31;
					xmlNode = xmlElement.SelectSingleNode(VH.A(196776));
					goto IL_0420;
					IL_0420:
					num2 = 32;
					UndoMaxCells = Conversions.ToInteger(xmlNode.Attributes[VH.A(196793)].Value);
					goto IL_0451;
					IL_0451:
					num2 = 33;
					UndoEnabled = Conversions.ToBoolean(xmlNode.Attributes[VH.A(94190)].Value);
					goto IL_0480;
					IL_0480:
					num2 = 34;
					UndoBorders = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(146542)).Attributes[VH.A(196808)].Value);
					goto IL_04c4;
					IL_04c4:
					num2 = 35;
					UndoFont = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(60635)).Attributes[VH.A(196808)].Value);
					goto IL_0506;
					IL_0506:
					num2 = 36;
					UndoFill = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(60234)).Attributes[VH.A(196808)].Value);
					goto IL_054a;
					IL_054a:
					num2 = 37;
					UndoAlignment = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(196819)).Attributes[VH.A(196808)].Value);
					goto IL_0590;
					IL_0590:
					xmlNode = null;
					goto IL_0593;
					IL_0593:
					num2 = 39;
					xmlNode2 = xmlElement.SelectSingleNode(VH.A(196838));
					goto IL_05ad;
					IL_05ad:
					num2 = 40;
					DependentsTimeout = Conversions.ToInteger(xmlNode2.SelectSingleNode(VH.A(196863)).InnerText);
					goto IL_05d3;
					IL_05d3:
					xmlNode2 = null;
					goto IL_05d6;
					IL_05d6:
					num2 = 42;
					num5 = clsUtilities.ParseDblInvariantCultureReplaceCommas(xmlElement.SelectSingleNode(VH.A(196898)).InnerText);
					if (!num5.HasValue)
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
						obj = null;
					}
					else
					{
						obj = num5.GetValueOrDefault() * 1000.0;
					}
					num5 = obj;
					TranslatorDelay = checked((int)Math.Round(num5.Value));
					goto IL_064e;
					IL_0686:
					num2 = 47;
					Check.CheckOutdatedLibraryContent = Conversions.ToBoolean(xmlElement.SelectSingleNode(Constants.XML_CHECK_OUTDATED_LIB_CONTENT).InnerText);
					goto IL_06aa;
					IL_06ad:
					xmlDocument = null;
					goto end_IL_0000_3;
					IL_06aa:
					xmlElement = null;
					goto IL_06ad;
					end_IL_0000_2:
					break;
				}
				num2 = 51;
				Forms.ErrorMessage(VH.A(196929));
				break;
				end_IL_0000:;
			}
			catch (object obj2) when (obj2 is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj2);
				try0000_dispatch = 1962;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	public void SaveSettings(XmlDocument xmlSettings)
	{
		Manage.Save(xmlSettings, true);
	}

	public static void SettingsExport()
	{
		Manage.Export();
	}

	public static bool SettingsImport()
	{
		Shortcuts.Remove();
		bool num = Manage.Import();
		if (num)
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
			A(VH.A(197111));
		}
		Shortcuts.Load();
		return num;
	}

	public static void SettingsReset()
	{
		Shortcuts.Remove();
		if (Manage.Reset())
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
			global::A.K.Settings.Reset();
			A(VH.A(197174));
		}
		Shortcuts.Load();
	}

	private static void A(string A)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 100:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_000f;
						case 4:
							goto IL_001b;
						case 5:
							goto IL_0022;
						case 6:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 7:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0022:
					num2 = 5;
					KH.A.Invalidate();
					break;
					IL_0007:
					num2 = 2;
					KH.A = null;
					goto IL_000f;
					IL_000f:
					num2 = 3;
					KH.A = new clsSettings();
					goto IL_001b;
					IL_001b:
					num2 = 4;
					Shortcuts.ResetShortcuts();
					goto IL_0022;
					end_IL_0000_2:
					break;
				}
				num2 = 6;
				Forms.SuccessMessage(A);
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 100;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
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
			ProjectData.ClearProjectError();
			return;
		}
	}
}
