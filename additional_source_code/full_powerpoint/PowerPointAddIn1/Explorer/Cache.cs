using System.Runtime.CompilerServices;
using System.Windows.Media;
using A;
using MacabacusMacros.Explorer;

namespace PowerPointAddIn1.Explorer;

public sealed class Cache
{
	[CompilerGenerated]
	private Geometry A;

	[CompilerGenerated]
	private Geometry B;

	[CompilerGenerated]
	private Geometry C;

	[CompilerGenerated]
	private Geometry D;

	[CompilerGenerated]
	private Geometry E;

	[CompilerGenerated]
	private Geometry F;

	[CompilerGenerated]
	private Geometry G;

	[CompilerGenerated]
	private Geometry H;

	[CompilerGenerated]
	private Geometry I;

	[CompilerGenerated]
	private Geometry J;

	[CompilerGenerated]
	private Geometry K;

	[CompilerGenerated]
	private Geometry L;

	[CompilerGenerated]
	private Geometry M;

	[CompilerGenerated]
	private Geometry N;

	[CompilerGenerated]
	private Geometry O;

	[CompilerGenerated]
	private Geometry P;

	[CompilerGenerated]
	private Geometry Q;

	[CompilerGenerated]
	private Geometry R;

	[CompilerGenerated]
	private Geometry S;

	[CompilerGenerated]
	private Geometry T;

	public Geometry GeoSlide
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	public Geometry GeoChart
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

	public Geometry GeoComment
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

	public Geometry GeoExcel
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

	public Geometry GeoWord
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

	public Geometry GeoHyperlink
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

	public Geometry GeoImage
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

	public Geometry GeoInk
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

	public Geometry GeoMedia
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

	public Geometry GeoNotes
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

	public Geometry GeoSmartArt
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

	public Geometry GeoTable
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

	public Geometry GeoTitle
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

	public Geometry GeoToC
	{
		[CompilerGenerated]
		get
		{
			return N;
		}
		[CompilerGenerated]
		set
		{
			N = value;
		}
	}

	public Geometry GeoFlysheet
	{
		[CompilerGenerated]
		get
		{
			return O;
		}
		[CompilerGenerated]
		set
		{
			O = value;
		}
	}

	public Geometry GeoFrontCover
	{
		[CompilerGenerated]
		get
		{
			return P;
		}
		[CompilerGenerated]
		set
		{
			P = value;
		}
	}

	public Geometry GeoBackCover
	{
		[CompilerGenerated]
		get
		{
			return Q;
		}
		[CompilerGenerated]
		set
		{
			Q = value;
		}
	}

	public Geometry GeoContact
	{
		[CompilerGenerated]
		get
		{
			return R;
		}
		[CompilerGenerated]
		set
		{
			R = value;
		}
	}

	public Geometry GeoLegal
	{
		[CompilerGenerated]
		get
		{
			return S;
		}
		[CompilerGenerated]
		set
		{
			S = value;
		}
	}

	public Geometry GeoBlank
	{
		[CompilerGenerated]
		get
		{
			return T;
		}
		[CompilerGenerated]
		set
		{
			T = value;
		}
	}

	public Cache()
	{
		GeoSlide = Geometry.Parse(AH.A(62867));
		GeoSlide.Freeze();
		GeoChart = Geometry.Parse(Constants.DATA_CHART_BASIC);
		GeoChart.Freeze();
		GeoComment = Geometry.Parse(Constants.DATA_COMMENT);
		GeoComment.Freeze();
		GeoExcel = Geometry.Parse(AH.A(106017));
		GeoExcel.Freeze();
		GeoWord = Geometry.Parse(AH.A(106475));
		GeoWord.Freeze();
		GeoHyperlink = Geometry.Parse(AH.A(106727));
		GeoHyperlink.Freeze();
		GeoImage = Geometry.Parse(AH.A(61793));
		GeoImage.Freeze();
		GeoInk = Geometry.Parse(AH.A(107383));
		GeoInk.Freeze();
		GeoMedia = Geometry.Parse(AH.A(107955));
		GeoMedia.Freeze();
		GeoNotes = Geometry.Parse(AH.A(108223));
		GeoNotes.Freeze();
		GeoSmartArt = Geometry.Parse(AH.A(108318));
		GeoSmartArt.Freeze();
		GeoTable = Geometry.Parse(AH.A(108594));
		GeoTable.Freeze();
		GeoTitle = Geometry.Parse(AH.A(109430));
		GeoTitle.Freeze();
		GeoToC = Geometry.Parse(AH.A(109610));
		GeoToC.Freeze();
		GeoFlysheet = Geometry.Parse(AH.A(110010));
		GeoFlysheet.Freeze();
		GeoFrontCover = Geometry.Parse(AH.A(110914));
		GeoFrontCover.Freeze();
		GeoBackCover = Geometry.Parse(AH.A(111094));
		GeoBackCover.Freeze();
		GeoContact = Geometry.Parse(AH.A(111274));
		GeoContact.Freeze();
		GeoLegal = Geometry.Parse(AH.A(112084));
		GeoLegal.Freeze();
		GeoBlank = Geometry.Parse(AH.A(112866));
		GeoBlank.Freeze();
	}
}
