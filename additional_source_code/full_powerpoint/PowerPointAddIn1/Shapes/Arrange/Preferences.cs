using System.Globalization;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Config.Settings;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes.Arrange;

public sealed class Preferences
{
	private readonly string m_A;

	private readonly string B;

	private readonly string C;

	private readonly string D;

	private readonly string E;

	private readonly string F;

	private readonly string G;

	private readonly string H;

	private readonly string I;

	private readonly string J;

	private readonly string K;

	private readonly string L;

	private readonly string M;

	[CompilerGenerated]
	private int m_A;

	[CompilerGenerated]
	private ScaleMode m_A;

	[CompilerGenerated]
	private bool m_A;

	[CompilerGenerated]
	private bool B;

	[CompilerGenerated]
	private float m_A;

	[CompilerGenerated]
	private float B;

	[CompilerGenerated]
	private float C;

	[CompilerGenerated]
	private Stretch m_A;

	[CompilerGenerated]
	private bool C;

	[CompilerGenerated]
	private bool D;

	[CompilerGenerated]
	private int B;

	[CompilerGenerated]
	private CircleAlign m_A;

	[CompilerGenerated]
	private int C;

	[CompilerGenerated]
	private bool E;

	[CompilerGenerated]
	private bool F;

	public int MaxShapesPerSlide
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

	public ScaleMode ScaleMode
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

	public bool ReorderBestFit
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

	public bool CenterShapes
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

	public float ContainerPadding
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

	public float MinColumnSpacing
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

	public float MinRowSpacing
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

	public Stretch StretchMode
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

	public bool StretchWidth
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

	public bool StretchHeight
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

	public int RotationAngle
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public CircleAlign CircleAlign
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

	public int CircleScale
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	public bool RotateShapes
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

	private bool IsMetric
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

	public Preferences()
	{
		this.m_A = AH.A(68510);
		this.B = AH.A(68551);
		this.C = AH.A(68586);
		this.D = AH.A(68605);
		this.E = AH.A(68634);
		this.F = AH.A(68659);
		G = AH.A(68692);
		H = AH.A(68725);
		I = AH.A(68752);
		J = AH.A(68775);
		K = AH.A(68798);
		L = AH.A(68821);
		M = AH.A(68848);
		XmlNode xmlNode = KG.A.SettingsXml.DocumentElement.SelectSingleNode(this.m_A);
		MaxShapesPerSlide = Conversions.ToInteger(xmlNode.SelectSingleNode(this.B).InnerText);
		ScaleMode = (ScaleMode)Conversions.ToInteger(xmlNode.SelectSingleNode(this.C).InnerText);
		ReorderBestFit = Conversions.ToBoolean(xmlNode.SelectSingleNode(this.D).InnerText);
		CenterShapes = Conversions.ToBoolean(xmlNode.SelectSingleNode(this.E).InnerText);
		ContainerPadding = Conversions.ToSingle(xmlNode.SelectSingleNode(this.F).InnerText);
		MinColumnSpacing = Conversions.ToSingle(xmlNode.SelectSingleNode(G).InnerText);
		MinRowSpacing = Conversions.ToSingle(xmlNode.SelectSingleNode(H).InnerText);
		StretchMode = (Stretch)Conversions.ToInteger(xmlNode.SelectSingleNode(I).InnerText);
		CircleAlign = (CircleAlign)Conversions.ToInteger(xmlNode.SelectSingleNode(J).InnerText);
		CircleScale = Conversions.ToInteger(xmlNode.SelectSingleNode(K).InnerText);
		RotationAngle = Conversions.ToInteger(xmlNode.SelectSingleNode(L).InnerText);
		RotateShapes = Conversions.ToBoolean(xmlNode.SelectSingleNode(M).InnerText);
		xmlNode = null;
		A(StretchMode);
		IsMetric = RegionInfo.CurrentRegion.IsMetric;
	}

	public void SaveMaxShapesPerSlide(int max)
	{
		MaxShapesPerSlide = max;
		A(this.B, Conversions.ToString(max));
	}

	public void SaveScaleMode(ScaleMode scale)
	{
		ScaleMode = scale;
		A(this.C, Conversions.ToString((int)scale));
	}

	public void SaveBestFit(bool blnBestFit)
	{
		ReorderBestFit = blnBestFit;
		A(this.D, blnBestFit.ToString());
	}

	public void SaveStretchMode(Stretch stretch)
	{
		StretchMode = stretch;
		A(stretch);
		A(I, Conversions.ToString((int)stretch));
	}

	public void SaveCenter(bool blnCenter)
	{
		CenterShapes = blnCenter;
		A(this.E, blnCenter.ToString());
	}

	public void SaveContainerPadding(float padding)
	{
		A(B: Conversions.ToString(ContainerPadding = A(padding)), A: this.F);
	}

	public void SaveMinColumnSpacing(float spacing)
	{
		A(B: Conversions.ToString(MinColumnSpacing = A(spacing)), A: G);
	}

	public void SaveMinRowSpacing(float spacing)
	{
		A(B: Conversions.ToString(MinRowSpacing = A(spacing)), A: H);
	}

	public void SaveCircleAlign(CircleAlign align)
	{
		CircleAlign = align;
		A(J, Conversions.ToString((int)align));
	}

	public void SaveCircleScale(int scale)
	{
		CircleScale = scale;
		A(K, Conversions.ToString(scale));
	}

	public void SaveRotationAngle(int angle)
	{
		RotationAngle = angle;
		A(L, Conversions.ToString(angle));
	}

	public void SaveRotateShapes(bool blnRotate)
	{
		RotateShapes = blnRotate;
		A(M, blnRotate.ToString());
	}

	private void A(string A, string B)
	{
		XmlDocument xml = Manage.GetXml(false);
		xml.DocumentElement.SelectSingleNode(this.m_A + AH.A(14622) + A).InnerText = B;
		_ = null;
		Manage.Save(xml, true);
	}

	public float ConvertFromPoints(float val)
	{
		if (!IsMetric)
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
					return clsPublish.PointsToInches(val);
				}
			}
		}
		return clsPublish.PointsToMillimeters(val);
	}

	private float A(float A)
	{
		if (!IsMetric)
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
					return clsPublish.InchesToPoints(A);
				}
			}
		}
		return clsPublish.MillimetersToPoints(A);
	}

	private void A(Stretch A)
	{
		switch (A)
		{
		case Stretch.WidthAndHeight:
			StretchHeight = true;
			StretchWidth = true;
			break;
		case Stretch.HeightOnly:
			StretchHeight = true;
			StretchWidth = false;
			break;
		case Stretch.WidthOnly:
			StretchHeight = false;
			StretchWidth = true;
			break;
		default:
			StretchHeight = false;
			StretchWidth = false;
			break;
		}
	}
}
