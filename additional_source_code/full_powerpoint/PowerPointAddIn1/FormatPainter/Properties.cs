using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.FormatPainter;

public sealed class Properties
{
	public struct ShapeProperties
	{
		public MsoShapeType Type;

		public MsoTriState HasTextFrame;

		public MsoTriState HasTable;

		public MsoTriState HasChart;

		public bool HasPicture;
	}

	public struct DecorationProperties
	{
		public MsoTriState Bold;

		public MsoTriState Italic;

		public MsoTextUnderlineType UnderlineStyle;

		public MsoTextStrike Strike;

		public MsoTriState StrikeThrough;

		public MsoTriState DoubleStrikeThrough;
	}

	public struct FontProperties
	{
		public float Size;

		public int ForeColor;

		public int BackColor;

		public int Highlight;

		public int UnderlineColor;

		public string Name;

		public DecorationProperties Decoration;
	}

	public struct FillProperties
	{
		public int BackColor;

		public int ForeColor;

		public MsoPatternType Pattern;

		public float Transparency;

		public MsoFillType Type;

		public MsoTriState Visible;

		public float GradientAngle;

		public MsoGradientColorType GradientColorType;

		public float GradientDegree;

		public GradientStops GradientStops;

		public MsoGradientStyle GradientStyle;

		public int GradientVariant;

		public MsoPresetGradientType PresetGradientType;
	}

	public struct LineProperties
	{
		public int BackColor;

		public int ForeColor;

		public MsoLineDashStyle DashStyle;

		public MsoLineStyle Style;

		public float Weight;

		public MsoTriState Visible;
	}

	public struct BulletProperties
	{
		public MsoNumberedBulletStyle Style;

		public MsoBulletType Type;

		public Font2 Font;

		public float RelativeSize;

		public int StartValue;

		public int Character;

		public MsoTriState UseTextColor;

		public MsoTriState UseTextFont;
	}

	public struct IndentProperties
	{
		public float FirstLineIndent;

		public float LeftIndent;

		public float RightIndent;
	}

	public struct SpacingProperties
	{
		public float SpaceAfter;

		public float SpaceBefore;

		public float SpaceWithin;

		public MsoTriState LineRuleWithin;
	}

	public struct TextBoxProperties
	{
		public Dictionary<int, BulletProperties> Bullets;

		public Dictionary<int, IndentProperties> Indents;

		public SpacingProperties LineSpacing;

		public MsoTextDirection TextDirection;

		public MsoAutoSize AutoSize;

		public float MarginTop;

		public float MarginBottom;

		public float MarginLeft;

		public float MarginRight;

		public MsoParagraphAlignment HorizontalAlignment;

		public MsoHorizontalAnchor HorizontalAnchor;

		public MsoVerticalAnchor VerticalAnchor;

		public MsoTriState WordWrap;

		public MsoTextOrientation Orientation;
	}

	public struct LayoutProperties
	{
		public float Height;

		public float Width;

		public float Left;

		public float Right;

		public float MidpointX;

		public float Top;

		public float Bottom;

		public float MidpointY;

		public float Rotation;

		public MsoTriState LockAspectRatio;
	}

	public struct AutoShapeProperties
	{
		public MsoAutoShapeType Type;

		public List<float> Adjustments;
	}

	public struct GlowProperties
	{
		public int Color;

		public float Radius;

		public float Transparency;
	}

	public struct ShapeShadowProperties
	{
		public float Blur;

		public int ForeColor;

		public MsoTriState Obscured;

		public float OffsetX;

		public float OffsetY;

		public MsoTriState RotateWithShape;

		public float Size;

		public MsoShadowStyle Style;

		public float Transparency;

		public MsoShadowType Type;

		public MsoTriState Visible;
	}

	public struct TextShadowProperties
	{
		public float Blur;

		public int ForeColor;

		public MsoTriState Obscured;

		public float OffsetX;

		public float OffsetY;

		public MsoTriState RotateWithShape;

		public float Size;

		public MsoShadowStyle Style;

		public float Transparency;

		public MsoShadowType Type;

		public MsoTriState Visible;
	}

	public struct ReflectionProperties
	{
		public float Blur;

		public float Offset;

		public float Size;

		public float Transparency;

		public MsoReflectionType Type;
	}

	public struct ThreeDProperties
	{
		public float BevelBottomDepth;

		public float BevelBottomInset;

		public MsoBevelType BevelBottomType;

		public float BevelTopDepth;

		public float BevelTopInset;

		public MsoBevelType BevelTopType;

		public int ContourColor;

		public float ContourWidth;

		public float Depth;

		public int ExtrusionColor;

		public MsoExtrusionColorType ExtrusionColorType;

		public float FieldOfView;

		public float LightAngle;

		public MsoTriState Perspective;

		public MsoPresetCamera PresetCamera;

		public MsoPresetExtrusionDirection PresetExtrusionDirection;

		public MsoLightRigType PresetLighting;

		public MsoPresetLightingDirection PresetLightingDirection;

		public MsoPresetLightingSoftness PresetLightingSoftness;

		public MsoPresetMaterial PresetMaterial;

		public MsoPresetThreeDFormat PresetThreeDFormat;

		public MsoTriState ProjectText;

		public float RotationX;

		public float RotationY;

		public float RotationZ;

		public MsoTriState Visible;

		public float Z;
	}

	[StructLayout(LayoutKind.Sequential, Size = 1)]
	public struct BevelProperties
	{
	}

	public struct SoftEdgeProperties
	{
		public float Radius;

		public MsoSoftEdgeType Type;
	}

	public struct ShapeEffectsProperties
	{
		public GlowProperties Glow;

		public ShapeShadowProperties Shadow;

		public ReflectionProperties Reflection;

		public ThreeDProperties ThreeD;

		public SoftEdgeProperties SoftEdge;

		public TextEffectFormat TextEffect;
	}

	public struct TextEffectsProperties
	{
		public GlowProperties Glow;

		public TextShadowProperties Shadow;

		public ReflectionProperties Reflection;

		public MsoSoftEdgeType SoftEdge;

		public ThreeDProperties ThreeD;

		public TextEffectFormat TextEffect;
	}

	public struct PlotAreaProperties
	{
		public float Top;

		public float Left;

		public float Height;

		public float Width;

		public double InsideTop;

		public double InsideLeft;

		public double InsideHeight;

		public double InsideWidth;

		public int ForeColor;

		public int BackColor;
	}

	public struct ChartProperties
	{
		public PlotAreaProperties PlotArea;
	}

	public struct PictureScale
	{
		public float ScaleHeight;

		public float ScaleWidth;
	}

	public struct PictureProperties
	{
		public float ScaleHeight;

		public float ScaleWidth;

		public float Brightness;

		public float Contrast;

		public Dictionary<MsoPictureEffectType, List<float>> PictureEffects;

		public float CropTop;

		public float CropBottom;

		public float CropLeft;

		public float CropRight;

		public float PictureOffsetY;

		public float PictureOffsetX;

		public float PictureHeight;

		public float PictureWidth;

		public float ShapeHeight;

		public float ShapeWidth;

		public float ShapeTop;

		public float ShapeLeft;

		public MsoPictureColorType ColorType;

		public int TransparencyColor;

		public MsoTriState TransparentBackground;
	}

	private ShapeProperties m_A;

	private FontProperties m_A;

	private FillProperties m_A;

	private LineProperties m_A;

	private TextBoxProperties m_A;

	private LayoutProperties m_A;

	private AutoShapeProperties m_A;

	private ChartProperties m_A;

	private PictureProperties m_A;

	private ShapeEffectsProperties m_A;

	private TextEffectsProperties m_A;

	public ShapeProperties Shape
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public FontProperties Font
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public FillProperties Fill
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public LineProperties Line
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public TextBoxProperties TextBox
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public LayoutProperties Layout
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public AutoShapeProperties AutoShape
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public ChartProperties Chart
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public PictureProperties Picture
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public ShapeEffectsProperties ShapeEffects
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public TextEffectsProperties TextEffects
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public Properties(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		try
		{
			Shape = new ShapeProperties
			{
				Type = shp.Type,
				HasTextFrame = shp.HasTextFrame,
				HasChart = shp.HasChart,
				HasTable = shp.HasTable,
				HasPicture = Images.HasPictureOrOLE(shp)
			};
			FillProperties fill = default(FillProperties);
			Microsoft.Office.Interop.PowerPoint.FillFormat fill2 = shp.Fill;
			fill.BackColor = fill2.BackColor.RGB;
			fill.ForeColor = fill2.ForeColor.RGB;
			fill.Pattern = fill2.Pattern;
			fill.Transparency = Math.Max(0f, fill2.Transparency);
			fill.Type = fill2.Type;
			fill.Visible = fill2.Visible;
			if (fill2.Type == MsoFillType.msoFillGradient)
			{
				fill.GradientAngle = fill2.GradientAngle;
				if (fill2.GradientColorType == MsoGradientColorType.msoGradientOneColor)
				{
					fill.GradientDegree = fill2.GradientDegree;
				}
				fill.GradientStops = fill2.GradientStops;
				fill.GradientStyle = fill2.GradientStyle;
				fill.GradientColorType = fill2.GradientColorType;
				fill.GradientVariant = fill2.GradientVariant;
				if (fill2.GradientColorType == MsoGradientColorType.msoGradientPresetColors)
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
					fill.PresetGradientType = fill2.PresetGradientType;
				}
			}
			fill2 = null;
			Fill = fill;
			LineProperties line = default(LineProperties);
			if (shp.HasTable == MsoTriState.msoFalse)
			{
				Microsoft.Office.Interop.PowerPoint.LineFormat line2 = shp.Line;
				line.BackColor = line2.BackColor.RGB;
				line.ForeColor = line2.ForeColor.RGB;
				line.DashStyle = line2.DashStyle;
				line.Style = line2.Style;
				line.Weight = line2.Weight;
				line.Visible = line2.Visible;
				line2 = null;
			}
			else
			{
				line.Visible = MsoTriState.msoFalse;
			}
			Line = line;
			if (shp.HasTextFrame == MsoTriState.msoTrue)
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
				FontProperties font = default(FontProperties);
				DecorationProperties decoration = default(DecorationProperties);
				Font2 font2 = shp.TextFrame2.TextRange.Font;
				decoration.Bold = font2.Bold;
				decoration.Italic = font2.Italic;
				decoration.UnderlineStyle = font2.UnderlineStyle;
				decoration.Strike = font2.Strike;
				decoration.StrikeThrough = font2.StrikeThrough;
				decoration.DoubleStrikeThrough = font2.DoubleStrikeThrough;
				font.Name = font2.Name;
				font.Size = font2.Size;
				font.ForeColor = font2.Fill.ForeColor.RGB;
				font.BackColor = font2.Fill.BackColor.RGB;
				font.Highlight = font2.Highlight.RGB;
				font.UnderlineColor = font2.UnderlineColor.RGB;
				font.Decoration = decoration;
				font2 = null;
				Font = font;
				TextBoxProperties textBox = default(TextBoxProperties);
				SpacingProperties lineSpacing = default(SpacingProperties);
				ParagraphFormat2 paragraphFormat = shp.TextFrame2.TextRange.ParagraphFormat;
				lineSpacing.SpaceAfter = paragraphFormat.SpaceAfter;
				lineSpacing.SpaceBefore = paragraphFormat.SpaceBefore;
				lineSpacing.SpaceWithin = paragraphFormat.SpaceWithin;
				lineSpacing.LineRuleWithin = paragraphFormat.LineRuleWithin;
				textBox.LineSpacing = lineSpacing;
				textBox.HorizontalAlignment = paragraphFormat.Alignment;
				textBox.TextDirection = paragraphFormat.TextDirection;
				paragraphFormat = null;
				Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = shp.TextFrame2;
				textBox.AutoSize = textFrame.AutoSize;
				textBox.MarginTop = textFrame.MarginTop;
				textBox.MarginBottom = textFrame.MarginBottom;
				textBox.MarginLeft = textFrame.MarginLeft;
				textBox.MarginRight = textFrame.MarginRight;
				textBox.HorizontalAnchor = textFrame.HorizontalAnchor;
				textBox.VerticalAnchor = textFrame.VerticalAnchor;
				textBox.Orientation = textFrame.Orientation;
				textBox.WordWrap = textFrame.WordWrap;
				textFrame = null;
				textBox.Bullets = GetBulletFormatting(shp);
				textBox.Indents = GetIndentFormatting(shp);
				TextBox = textBox;
				TextEffectsProperties textEffects = default(TextEffectsProperties);
				A(shp.TextFrame2.ThreeD);
				Font2 font3 = shp.TextFrame2.TextRange.Font;
				GlowProperties glow = A(font3.Glow);
				TextShadowProperties shadow = A(font3.Shadow);
				ReflectionProperties reflection = A(font3.Reflection);
				textEffects.SoftEdge = font3.SoftEdgeFormat;
				font3 = null;
				textEffects.Glow = glow;
				textEffects.Shadow = shadow;
				textEffects.Reflection = reflection;
				TextEffects = textEffects;
			}
			LayoutProperties layout = default(LayoutProperties);
			Microsoft.Office.Interop.PowerPoint.Shape shape = shp;
			layout.Height = shp.Height;
			layout.Width = shp.Width;
			layout.Left = shape.Left;
			layout.Right = shape.Left + shape.Width;
			layout.MidpointX = shape.Left + shape.Width / 2f;
			layout.Top = shape.Top;
			layout.Bottom = shape.Top + shape.Height;
			layout.MidpointY = shape.Top + shape.Height / 2f;
			try
			{
				layout.Rotation = shape.Rotation;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				layout.Rotation = 0f;
				ProjectData.ClearProjectError();
			}
			layout.LockAspectRatio = shape.LockAspectRatio;
			shape = null;
			Layout = layout;
			ShapeEffectsProperties shapeEffects = default(ShapeEffectsProperties);
			SoftEdgeProperties softEdge = default(SoftEdgeProperties);
			Microsoft.Office.Interop.PowerPoint.Shape shape2 = shp;
			GlowProperties glow2 = default(GlowProperties);
			if (shp.HasTable == MsoTriState.msoFalse)
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
				glow2 = A(shape2.Glow);
				SoftEdgeFormat softEdge2 = shape2.SoftEdge;
				softEdge.Radius = softEdge2.Radius;
				softEdge.Type = softEdge2.Type;
				softEdge2 = null;
			}
			else
			{
				softEdge.Type = MsoSoftEdgeType.msoSoftEdgeTypeNone;
			}
			ReflectionProperties reflection2;
			if (shp.HasChart == MsoTriState.msoFalse)
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
				reflection2 = A(shape2.Reflection);
			}
			else
			{
				reflection2 = new ReflectionProperties
				{
					Type = MsoReflectionType.msoReflectionTypeNone
				};
			}
			ShapeShadowProperties shadow2 = A(shape2.Shadow);
			ThreeDProperties threeD = A(shape2.ThreeD);
			try
			{
				shapeEffects.TextEffect = shape2.TextEffect;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			shape2 = null;
			shapeEffects.Glow = glow2;
			shapeEffects.Shadow = shadow2;
			shapeEffects.Reflection = reflection2;
			shapeEffects.ThreeD = threeD;
			shapeEffects.SoftEdge = softEdge;
			ShapeEffects = shapeEffects;
			AutoShapeProperties autoShape = default(AutoShapeProperties);
			Microsoft.Office.Interop.PowerPoint.Shape shape3 = shp;
			if (shape3.Type == MsoShapeType.msoAutoShape)
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
				List<float> list = new List<float>();
				autoShape.Type = shape3.AutoShapeType;
				int count = shape3.Adjustments.Count;
				for (int i = 1; i <= count; i = checked(i + 1))
				{
					list.Add(shape3.Adjustments[i]);
				}
				autoShape.Adjustments = list;
				list = null;
			}
			shape3 = null;
			AutoShape = autoShape;
			PictureProperties picture = default(PictureProperties);
			if (Images.HasPictureOrOLE(shp))
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
				Microsoft.Office.Interop.PowerPoint.Shape shape4 = shp;
				picture.Brightness = shape4.PictureFormat.Brightness;
				picture.Contrast = shape4.PictureFormat.Contrast;
				picture.PictureEffects = new Dictionary<MsoPictureEffectType, List<float>>();
				try
				{
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = shp.Fill.PictureEffects.GetEnumerator();
						IEnumerator enumerator2 = default(IEnumerator);
						while (enumerator.MoveNext())
						{
							PictureEffect pictureEffect = (PictureEffect)enumerator.Current;
							List<float> list2 = new List<float>();
							try
							{
								enumerator2 = pictureEffect.EffectParameters.GetEnumerator();
								while (enumerator2.MoveNext())
								{
									EffectParameter effectParameter = (EffectParameter)enumerator2.Current;
									list2.Add(Conversions.ToSingle(effectParameter.Value));
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										goto end_IL_08c9;
									}
									continue;
									end_IL_08c9:
									break;
								}
							}
							finally
							{
								if (enumerator2 is IDisposable)
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											continue;
										}
										(enumerator2 as IDisposable).Dispose();
										break;
									}
								}
							}
							picture.PictureEffects.Add(pictureEffect.Type, list2);
							list2 = null;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_091d;
							}
							continue;
							end_IL_091d:
							break;
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
								(enumerator as IDisposable).Dispose();
								break;
							}
						}
					}
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
				PictureScale pictureScale = A(shp);
				picture.ScaleHeight = pictureScale.ScaleHeight;
				picture.ScaleWidth = pictureScale.ScaleWidth;
				PictureFormat pictureFormat = shape4.PictureFormat;
				try
				{
					picture.PictureOffsetY = pictureFormat.Crop.PictureOffsetY;
					picture.PictureOffsetX = pictureFormat.Crop.PictureOffsetX;
					picture.PictureHeight = pictureFormat.Crop.PictureHeight;
					picture.PictureWidth = pictureFormat.Crop.PictureWidth;
					picture.ShapeHeight = pictureFormat.Crop.ShapeHeight;
					picture.ShapeWidth = pictureFormat.Crop.ShapeWidth;
					picture.ShapeTop = pictureFormat.Crop.ShapeTop;
					picture.ShapeLeft = pictureFormat.Crop.ShapeLeft;
					picture.CropBottom = pictureFormat.CropBottom;
					picture.CropLeft = pictureFormat.CropLeft;
					picture.CropRight = pictureFormat.CropRight;
					picture.CropTop = pictureFormat.CropTop;
				}
				catch (Exception ex7)
				{
					ProjectData.SetProjectError(ex7);
					Exception ex8 = ex7;
					ProjectData.ClearProjectError();
				}
				pictureFormat = null;
				shape4 = null;
			}
			Picture = picture;
		}
		catch (Exception ex9)
		{
			ProjectData.SetProjectError(ex9);
			Exception ex10 = ex9;
			Forms.ErrorMessage(AH.A(139622) + ex10.Message);
			clsReporting.LogException(ex10);
			ProjectData.ClearProjectError();
		}
	}

	public static Dictionary<int, BulletProperties> GetBulletFormatting(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		Dictionary<int, BulletProperties> dictionary = new Dictionary<int, BulletProperties>();
		foreach (TextRange2 item in shp.TextFrame2.TextRange.get_Paragraphs(-1, -1))
		{
			ParagraphFormat2 paragraphFormat = item.ParagraphFormat;
			if (paragraphFormat.Bullet.Type != MsoBulletType.msoBulletNone && !dictionary.ContainsKey(paragraphFormat.IndentLevel))
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
				BulletProperties value = default(BulletProperties);
				BulletFormat2 bullet = paragraphFormat.Bullet;
				value.Style = bullet.Style;
				value.Type = bullet.Type;
				value.Font = bullet.Font;
				value.RelativeSize = bullet.RelativeSize;
				value.StartValue = bullet.StartValue;
				value.Character = bullet.Character;
				value.UseTextColor = bullet.UseTextColor;
				value.UseTextFont = bullet.UseTextFont;
				bullet = null;
				dictionary.Add(paragraphFormat.IndentLevel, value);
			}
			paragraphFormat = null;
		}
		return dictionary;
	}

	public static Dictionary<int, IndentProperties> GetIndentFormatting(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		Dictionary<int, IndentProperties> dictionary = new Dictionary<int, IndentProperties>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = shp.TextFrame2.TextRange.get_Paragraphs(-1, -1).GetEnumerator();
			while (enumerator.MoveNext())
			{
				ParagraphFormat2 paragraphFormat = ((TextRange2)enumerator.Current).ParagraphFormat;
				if (!dictionary.ContainsKey(paragraphFormat.IndentLevel))
				{
					IndentProperties value = default(IndentProperties);
					value.LeftIndent = paragraphFormat.LeftIndent;
					value.FirstLineIndent = paragraphFormat.FirstLineIndent;
					value.LeftIndent = paragraphFormat.LeftIndent;
					value.RightIndent = paragraphFormat.RightIndent;
					dictionary.Add(paragraphFormat.IndentLevel, value);
				}
				paragraphFormat = null;
			}
			return dictionary;
		}
		finally
		{
			if (enumerator is IDisposable)
			{
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
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	private ThreeDProperties A(Microsoft.Office.Interop.PowerPoint.ThreeDFormat A)
	{
		ThreeDProperties result = default(ThreeDProperties);
		Microsoft.Office.Interop.PowerPoint.ThreeDFormat threeDFormat = A;
		result.BevelBottomDepth = threeDFormat.BevelBottomDepth;
		result.BevelBottomInset = threeDFormat.BevelBottomInset;
		result.BevelBottomType = threeDFormat.BevelBottomType;
		result.BevelTopDepth = threeDFormat.BevelTopDepth;
		result.BevelTopInset = threeDFormat.BevelTopInset;
		result.BevelTopType = threeDFormat.BevelTopType;
		result.ContourColor = threeDFormat.ContourColor.RGB;
		result.ContourWidth = threeDFormat.ContourWidth;
		result.Depth = threeDFormat.Depth;
		result.ExtrusionColor = threeDFormat.ExtrusionColor.RGB;
		result.ExtrusionColorType = threeDFormat.ExtrusionColorType;
		result.FieldOfView = threeDFormat.FieldOfView;
		result.LightAngle = threeDFormat.LightAngle;
		result.Perspective = threeDFormat.Perspective;
		result.PresetCamera = threeDFormat.PresetCamera;
		result.PresetExtrusionDirection = threeDFormat.PresetExtrusionDirection;
		result.PresetLighting = threeDFormat.PresetLighting;
		result.PresetLightingDirection = threeDFormat.PresetLightingDirection;
		result.PresetLightingSoftness = threeDFormat.PresetLightingSoftness;
		result.PresetMaterial = threeDFormat.PresetMaterial;
		result.PresetThreeDFormat = threeDFormat.PresetThreeDFormat;
		result.ProjectText = threeDFormat.ProjectText;
		result.RotationX = threeDFormat.RotationX;
		result.RotationY = threeDFormat.RotationY;
		result.RotationZ = threeDFormat.RotationZ;
		result.Visible = threeDFormat.Visible;
		result.Z = threeDFormat.Z;
		threeDFormat = null;
		return result;
	}

	private ShapeShadowProperties A(Microsoft.Office.Interop.PowerPoint.ShadowFormat A)
	{
		ShapeShadowProperties result = default(ShapeShadowProperties);
		Microsoft.Office.Interop.PowerPoint.ShadowFormat shadowFormat = A;
		result.Blur = shadowFormat.Blur;
		result.ForeColor = shadowFormat.ForeColor.RGB;
		result.Obscured = shadowFormat.Obscured;
		result.OffsetX = shadowFormat.OffsetX;
		result.OffsetY = shadowFormat.OffsetY;
		result.RotateWithShape = shadowFormat.RotateWithShape;
		result.Size = shadowFormat.Size;
		result.Style = shadowFormat.Style;
		result.Transparency = shadowFormat.Transparency;
		result.Type = shadowFormat.Type;
		result.Visible = shadowFormat.Visible;
		shadowFormat = null;
		return result;
	}

	private TextShadowProperties A(Microsoft.Office.Core.ShadowFormat A)
	{
		TextShadowProperties result = default(TextShadowProperties);
		Microsoft.Office.Core.ShadowFormat shadowFormat = A;
		result.Blur = shadowFormat.Blur;
		result.ForeColor = shadowFormat.ForeColor.RGB;
		result.Obscured = shadowFormat.Obscured;
		result.OffsetX = shadowFormat.OffsetX;
		result.OffsetY = shadowFormat.OffsetY;
		result.RotateWithShape = shadowFormat.RotateWithShape;
		result.Size = shadowFormat.Size;
		result.Style = shadowFormat.Style;
		result.Transparency = shadowFormat.Transparency;
		result.Type = shadowFormat.Type;
		result.Visible = shadowFormat.Visible;
		shadowFormat = null;
		return result;
	}

	private ReflectionProperties A(ReflectionFormat A)
	{
		ReflectionProperties result = default(ReflectionProperties);
		ReflectionFormat reflectionFormat = A;
		result.Blur = reflectionFormat.Blur;
		result.Offset = reflectionFormat.Offset;
		result.Size = reflectionFormat.Size;
		result.Transparency = reflectionFormat.Transparency;
		result.Type = reflectionFormat.Type;
		reflectionFormat = null;
		return result;
	}

	private GlowProperties A(GlowFormat A)
	{
		GlowProperties result = default(GlowProperties);
		GlowFormat glowFormat = A;
		result.Color = glowFormat.Color.RGB;
		result.Radius = glowFormat.Radius;
		result.Transparency = glowFormat.Transparency;
		glowFormat = null;
		return result;
	}

	private float A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		float result = 0f;
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		try
		{
			result = shape.Rotation;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (A.Type == MsoShapeType.msoAutoShape)
		{
			if (shape.AutoShapeType == MsoAutoShapeType.msoShapeMixed)
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
				if (A.Line.EndArrowheadStyle != MsoArrowheadStyle.msoArrowheadTriangle)
				{
					if (A.Line.BeginArrowheadStyle != MsoArrowheadStyle.msoArrowheadTriangle)
					{
						goto IL_00a1;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
				}
				result = B(A);
			}
		}
		else if (shape.Type == MsoShapeType.msoLine)
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
			result = B(A);
		}
		goto IL_00a1;
		IL_00a1:
		shape = null;
		return result;
	}

	private static float B(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		float num = (float)(Math.Atan2(shape.Height, shape.Width) * 57.2957795);
		float result;
		if (shape.VerticalFlip == MsoTriState.msoTrue)
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
			if (shape.HorizontalFlip == MsoTriState.msoTrue)
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
				result = num;
			}
			else
			{
				result = 360f - num;
			}
		}
		else if (shape.HorizontalFlip == MsoTriState.msoTrue)
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
			result = 360f - num;
		}
		else
		{
			result = num;
		}
		shape = null;
		return result;
	}

	private PictureScale A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		PictureScale result = default(PictureScale);
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		MsoTriState lockAspectRatio = shape.LockAspectRatio;
		shape.LockAspectRatio = MsoTriState.msoFalse;
		float width = shape.Width;
		shape.ScaleWidth(1f, MsoTriState.msoTrue);
		float width2 = shape.Width;
		float num = width / width2;
		shape.ScaleWidth(num, MsoTriState.msoTrue);
		float height = shape.Height;
		shape.ScaleHeight(1f, MsoTriState.msoTrue);
		float height2 = shape.Height;
		float num2 = height / height2;
		shape.ScaleHeight(num2, MsoTriState.msoTrue);
		if (num == num2)
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
			shape.LockAspectRatio = lockAspectRatio;
		}
		shape = null;
		result.ScaleWidth = num;
		result.ScaleHeight = num2;
		return result;
	}
}
