using System.Runtime.CompilerServices;

namespace PowerPointAddIn1.FormatPainter;

public sealed class Options
{
	public sealed class myFont
	{
		[CompilerGenerated]
		private bool A;

		[CompilerGenerated]
		private bool B;

		[CompilerGenerated]
		private bool C;

		[CompilerGenerated]
		private bool D;

		public bool Size
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

		public bool Color
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

		public bool Name
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

		public bool Decoration
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

		public myFont()
		{
			Size = false;
			Color = false;
			Name = false;
			Decoration = false;
		}
	}

	public sealed class myFill
	{
		[CompilerGenerated]
		private bool A;

		[CompilerGenerated]
		private bool B;

		[CompilerGenerated]
		private bool C;

		public bool Color
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

		public bool Type
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

		public bool Transparency
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

		public myFill()
		{
			Color = false;
			Type = false;
			Transparency = false;
		}
	}

	public sealed class myLine
	{
		[CompilerGenerated]
		private bool A;

		[CompilerGenerated]
		private bool B;

		[CompilerGenerated]
		private bool C;

		public bool Color
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

		public bool Style
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

		public bool Weight
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

		public myLine()
		{
			Color = false;
			Style = false;
			Weight = false;
		}
	}

	public sealed class myTextBox
	{
		[CompilerGenerated]
		private bool A;

		[CompilerGenerated]
		private bool B;

		[CompilerGenerated]
		private bool C;

		[CompilerGenerated]
		private bool D;

		[CompilerGenerated]
		private bool E;

		[CompilerGenerated]
		private bool F;

		[CompilerGenerated]
		private bool G;

		[CompilerGenerated]
		private bool H;

		[CompilerGenerated]
		private bool I;

		public bool Bullets
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

		public bool Indents
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

		public bool LineSpacing
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

		public bool Margins
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

		public bool AutoSize
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

		public bool WordWrap
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

		public bool HorizontalAlignment
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

		public bool VerticalAlignment
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

		public bool Orientation
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

		public myTextBox()
		{
			Bullets = false;
			Indents = false;
			LineSpacing = false;
			Margins = false;
			AutoSize = false;
			WordWrap = false;
			HorizontalAlignment = false;
			VerticalAlignment = false;
			Orientation = false;
		}
	}

	public sealed class myLayout
	{
		[CompilerGenerated]
		private bool A;

		[CompilerGenerated]
		private bool B;

		[CompilerGenerated]
		private bool C;

		[CompilerGenerated]
		private bool D;

		[CompilerGenerated]
		private bool E;

		[CompilerGenerated]
		private bool F;

		[CompilerGenerated]
		private bool G;

		[CompilerGenerated]
		private bool H;

		[CompilerGenerated]
		private bool I;

		[CompilerGenerated]
		private bool J;

		public bool Width
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

		public bool Height
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

		public bool Top
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

		public bool Bottom
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

		public bool MidpointY
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

		public bool Left
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

		public bool Right
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

		public bool MidpointX
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

		public bool Rotation
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

		public bool LockAspectRatio
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

		public myLayout()
		{
			Width = false;
			Height = false;
			Top = false;
			Bottom = false;
			MidpointY = false;
			Left = false;
			Right = false;
			MidpointX = false;
			Rotation = false;
			LockAspectRatio = false;
		}
	}

	public sealed class myShapeEffects
	{
		[CompilerGenerated]
		private bool A;

		[CompilerGenerated]
		private bool B;

		[CompilerGenerated]
		private bool C;

		[CompilerGenerated]
		private bool D;

		[CompilerGenerated]
		private bool E;

		public bool Shadow
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

		public bool Glow
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

		public bool Reflection
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

		public bool SoftEdge
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

		public bool ThreeD
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

		public myShapeEffects()
		{
			Shadow = false;
			Glow = false;
			Reflection = false;
			SoftEdge = false;
			ThreeD = false;
		}
	}

	public sealed class myTextEffects
	{
		[CompilerGenerated]
		private bool A;

		[CompilerGenerated]
		private bool B;

		[CompilerGenerated]
		private bool C;

		[CompilerGenerated]
		private bool D;

		[CompilerGenerated]
		private bool E;

		public bool Shadow
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

		public bool Glow
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

		public bool Reflection
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

		public bool SoftEdge
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

		public bool ThreeD
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

		public myTextEffects()
		{
			Shadow = false;
			Glow = false;
			Reflection = false;
			SoftEdge = false;
			ThreeD = false;
		}
	}

	public sealed class myPicture
	{
		[CompilerGenerated]
		private bool A;

		[CompilerGenerated]
		private bool B;

		[CompilerGenerated]
		private bool C;

		[CompilerGenerated]
		private bool D;

		[CompilerGenerated]
		private bool E;

		[CompilerGenerated]
		private bool F;

		[CompilerGenerated]
		private bool G;

		[CompilerGenerated]
		private bool H;

		[CompilerGenerated]
		private bool I;

		public bool ScaleHeight
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

		public bool ScaleWidth
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

		public bool Crop
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

		public bool Brightness
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

		public bool Contrast
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

		public bool Sharpness
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

		public bool Saturation
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

		public bool Temperature
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

		public bool Transparency
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

		public myPicture()
		{
			ScaleHeight = false;
			ScaleWidth = false;
			Crop = false;
			Brightness = false;
			Contrast = false;
			Sharpness = false;
			Saturation = false;
			Temperature = false;
			Transparency = false;
		}
	}

	public sealed class myAutoShape
	{
		[CompilerGenerated]
		private bool A;

		[CompilerGenerated]
		private bool B;

		public bool Type
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

		public bool Adjustments
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

		public myAutoShape()
		{
			Type = false;
			Adjustments = false;
		}
	}

	[CompilerGenerated]
	private myFont A;

	[CompilerGenerated]
	private myFill A;

	[CompilerGenerated]
	private myLine A;

	[CompilerGenerated]
	private myTextBox A;

	[CompilerGenerated]
	private myLayout A;

	[CompilerGenerated]
	private myShapeEffects A;

	[CompilerGenerated]
	private myTextEffects A;

	[CompilerGenerated]
	private myAutoShape A;

	[CompilerGenerated]
	private myPicture A;

	public myFont Font
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public myFill Fill
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public myLine Line
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public myTextBox TextBox
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public myLayout Layout
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public myShapeEffects ShapeEffects
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public myTextEffects TextEffects
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public myAutoShape AutoShape
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public myPicture Picture
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

	public Options()
	{
		Font = new myFont();
		Fill = new myFill();
		Line = new myLine();
		TextBox = new myTextBox();
		Layout = new myLayout();
		ShapeEffects = new myShapeEffects();
		TextEffects = new myTextEffects();
		AutoShape = new myAutoShape();
		Picture = new myPicture();
	}
}
