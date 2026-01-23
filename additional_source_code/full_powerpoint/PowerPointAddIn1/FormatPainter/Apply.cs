using System;
using System.Collections;
using System.Collections.Generic;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.FormatPainter;

public sealed class Apply
{
	public static void ToSelection(Properties properties, Options options)
	{
		if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			IEnumerator enumerator2 = default(IEnumerator);
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
				Application application = NG.A.Application;
				Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange;
				try
				{
					shapeRange = Base.SelectedShapes(application.ActiveWindow.Selection);
					try
					{
						application.StartNewUndoEntry();
						try
						{
							enumerator = shapeRange.GetEnumerator();
							while (enumerator.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
								if (!options.Fill.Color && !options.Fill.Type)
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
									if (!options.Fill.Transparency)
									{
										goto IL_0167;
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											continue;
										}
										break;
									}
								}
								if (shape.HasTable == MsoTriState.msoFalse)
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
									A(shape, properties, options);
								}
								else
								{
									Table table = shape.Table;
									int count = table.Rows.Count;
									int count2 = table.Columns.Count;
									int num = count;
									for (int i = 1; i <= num; i++)
									{
										int num2 = count2;
										for (int j = 1; j <= num2; j++)
										{
											Cell cell = table.Cell(i, j);
											if (cell.Selected)
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
												A(cell.Shape, properties, options);
											}
											cell = null;
										}
										while (true)
										{
											switch (3)
											{
											case 0:
												break;
											default:
												goto end_IL_014e;
											}
											continue;
											end_IL_014e:
											break;
										}
									}
									table = null;
								}
								goto IL_0167;
								IL_0f82:
								Properties.PictureProperties picture;
								PictureFormat pictureFormat;
								if (options.Picture.Transparency)
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
									pictureFormat.TransparencyColor = picture.TransparencyColor;
									pictureFormat.TransparentBackground = picture.TransparentBackground;
								}
								picture = default(Properties.PictureProperties);
								pictureFormat = null;
								continue;
								IL_02ae:
								if (shape.HasTextFrame == MsoTriState.msoTrue)
								{
									A(shape.TextFrame2.TextRange, properties, options);
									A(shape.TextFrame2, properties, options);
								}
								else if (shape.HasTable == MsoTriState.msoTrue)
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
									Table table2 = shape.Table;
									int count3 = table2.Rows.Count;
									int count4 = table2.Columns.Count;
									int num3 = count3;
									for (int k = 1; k <= num3; k++)
									{
										int num4 = count4;
										for (int l = 1; l <= num4; l++)
										{
											Cell cell2 = table2.Cell(k, l);
											if (cell2.Selected)
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
												if (cell2.Shape.HasTextFrame == MsoTriState.msoTrue)
												{
													A(cell2.Shape.TextFrame2.TextRange, properties, options);
													A(cell2.Shape.TextFrame2, properties, options);
												}
											}
											cell2 = null;
										}
										while (true)
										{
											switch (7)
											{
											case 0:
												break;
											default:
												goto end_IL_03ab;
											}
											continue;
											end_IL_03ab:
											break;
										}
									}
									table2 = null;
								}
								else if (shape.HasSmartArt == MsoTriState.msoTrue)
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
									try
									{
										enumerator2 = shape.SmartArt.AllNodes.GetEnumerator();
										while (enumerator2.MoveNext())
										{
											SmartArtNode obj = (SmartArtNode)enumerator2.Current;
											A(obj.TextFrame2.TextRange, properties, options);
											Microsoft.Office.Core.TextFrame2 textFrame = obj.TextFrame2;
											Properties.TextBoxProperties textBox = properties.TextBox;
											if (options.TextBox.Margins)
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
												textFrame.MarginTop = textBox.MarginTop;
												textFrame.MarginBottom = textBox.MarginBottom;
												textFrame.MarginLeft = textBox.MarginLeft;
												textFrame.MarginRight = textBox.MarginRight;
											}
											if (options.TextBox.AutoSize)
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
												textFrame.AutoSize = textBox.AutoSize;
											}
											if (options.TextBox.WordWrap)
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
												textFrame.WordWrap = textBox.WordWrap;
											}
											if (options.TextBox.HorizontalAlignment)
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
												textFrame.HorizontalAnchor = textBox.HorizontalAnchor;
											}
											if (options.TextBox.VerticalAlignment)
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
												textFrame.VerticalAnchor = textBox.VerticalAnchor;
											}
											if (options.TextBox.Orientation)
											{
												textFrame.Orientation = textBox.Orientation;
											}
											textBox = default(Properties.TextBoxProperties);
											Properties.TextEffectsProperties textEffects = properties.TextEffects;
											if (options.TextEffects.ThreeD)
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
												A((Microsoft.Office.Interop.PowerPoint.ThreeDFormat)textFrame.ThreeD, textEffects.ThreeD);
											}
											textEffects = default(Properties.TextEffectsProperties);
										}
										while (true)
										{
											switch (1)
											{
											case 0:
												break;
											default:
												goto end_IL_0597;
											}
											continue;
											end_IL_0597:
											break;
										}
									}
									finally
									{
										if (enumerator2 is IDisposable)
										{
											while (true)
											{
												switch (4)
												{
												case 0:
													continue;
												}
												(enumerator2 as IDisposable).Dispose();
												break;
											}
										}
									}
								}
								MsoTriState lockAspectRatio = shape.LockAspectRatio;
								shape.LockAspectRatio = MsoTriState.msoFalse;
								Properties.LayoutProperties layout = properties.Layout;
								if (options.Layout.Height)
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
									shape.Height = layout.Height;
								}
								if (options.Layout.Width)
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
									shape.Width = layout.Width;
								}
								if (options.Layout.LockAspectRatio)
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
									shape.LockAspectRatio = layout.LockAspectRatio;
								}
								else
								{
									shape.LockAspectRatio = lockAspectRatio;
								}
								if (options.Layout.Top)
								{
									shape.Top = layout.Top;
								}
								else if (options.Layout.Bottom)
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
									shape.Top = layout.Bottom - shape.Height;
								}
								else if (options.Layout.MidpointY)
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
									shape.Top = layout.MidpointY - shape.Height / 2f;
								}
								if (options.Layout.Left)
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
									shape.Left = layout.Left;
								}
								else if (options.Layout.Right)
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
									shape.Left = layout.Right - shape.Width;
								}
								else if (options.Layout.MidpointX)
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
									shape.Left = layout.MidpointX - shape.Width / 2f;
								}
								if (options.Layout.Rotation)
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
									try
									{
										shape.Rotation = layout.Rotation;
									}
									catch (Exception ex)
									{
										ProjectData.SetProjectError(ex);
										Exception ex2 = ex;
										ProjectData.ClearProjectError();
									}
								}
								Properties.ShapeEffectsProperties shapeEffects = properties.ShapeEffects;
								if (options.ShapeEffects.Glow)
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
									if (shape.HasTable == MsoTriState.msoFalse)
									{
										A(shape.Glow, shapeEffects.Glow);
									}
								}
								if (options.ShapeEffects.Shadow)
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
									A(shape.Shadow, shapeEffects.Shadow);
								}
								if (options.ShapeEffects.Reflection)
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
									if (shape.HasChart == MsoTriState.msoFalse)
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
										A(shape.Reflection, shapeEffects.Reflection);
									}
								}
								if (options.ShapeEffects.ThreeD)
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
									A(shape.ThreeD, shapeEffects.ThreeD);
								}
								if (options.ShapeEffects.SoftEdge && shape.HasTable == MsoTriState.msoFalse)
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
									try
									{
										Properties.SoftEdgeProperties softEdge = shapeEffects.SoftEdge;
										shape.SoftEdge.Type = softEdge.Type;
										if (softEdge.Type != MsoSoftEdgeType.msoSoftEdgeTypeNone)
										{
											shape.SoftEdge.Radius = softEdge.Radius;
										}
									}
									catch (Exception ex3)
									{
										ProjectData.SetProjectError(ex3);
										Exception ex4 = ex3;
										ProjectData.ClearProjectError();
									}
								}
								shapeEffects = default(Properties.ShapeEffectsProperties);
								if (shape.Type == MsoShapeType.msoAutoShape)
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
									Properties.AutoShapeProperties autoShape = properties.AutoShape;
									if (options.AutoShape.Type)
									{
										shape.AutoShapeType = autoShape.Type;
									}
									if (options.AutoShape.Adjustments)
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
										int num5 = autoShape.Adjustments.Count - 1;
										for (int m = 0; m <= num5; m++)
										{
											shape.Adjustments[m + 1] = autoShape.Adjustments[m];
										}
										while (true)
										{
											switch (2)
											{
											case 0:
												continue;
											}
											break;
										}
									}
									autoShape = default(Properties.AutoShapeProperties);
								}
								if (!Images.HasPictureOrOLE(shape))
								{
									continue;
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									break;
								}
								pictureFormat = shape.PictureFormat;
								picture = properties.Picture;
								if (!options.Picture.Sharpness)
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
									if (!options.Picture.Brightness && !options.Picture.Contrast)
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
										if (!options.Picture.Saturation)
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
											if (!options.Picture.Temperature)
											{
												goto IL_0e84;
											}
										}
									}
								}
								for (int n = shape.Fill.PictureEffects.Count; n >= 1; n += -1)
								{
									PictureEffect pictureEffect = shape.Fill.PictureEffects[n];
									MsoPictureEffectType type = pictureEffect.Type;
									if (type <= MsoPictureEffectType.msoEffectColorTemperature)
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
										switch (type)
										{
										default:
											while (true)
											{
												switch (5)
												{
												case 0:
													continue;
												}
												break;
											}
											break;
										case MsoPictureEffectType.msoEffectBrightnessContrast:
											if (!options.Picture.Brightness)
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
												if (!options.Picture.Contrast)
												{
													break;
												}
												while (true)
												{
													switch (1)
													{
													case 0:
														continue;
													}
													break;
												}
											}
											pictureEffect.Delete();
											break;
										case MsoPictureEffectType.msoEffectColorTemperature:
											if (!options.Picture.Temperature)
											{
												break;
											}
											while (true)
											{
												switch (6)
												{
												case 0:
													continue;
												}
												break;
											}
											pictureEffect.Delete();
											break;
										}
									}
									else if (type != MsoPictureEffectType.msoEffectSaturation)
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
										if (type == MsoPictureEffectType.msoEffectSharpenSoften && options.Picture.Sharpness)
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
											pictureEffect.Delete();
										}
									}
									else if (options.Picture.Saturation)
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
										pictureEffect.Delete();
									}
									pictureEffect = null;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									break;
								}
								bool flag = false;
								using (Dictionary<MsoPictureEffectType, List<float>>.Enumerator enumerator3 = picture.PictureEffects.GetEnumerator())
								{
									while (enumerator3.MoveNext())
									{
										KeyValuePair<MsoPictureEffectType, List<float>> current = enumerator3.Current;
										MsoPictureEffectType key = current.Key;
										if (key <= MsoPictureEffectType.msoEffectColorTemperature)
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
											if (key != MsoPictureEffectType.msoEffectBrightnessContrast)
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
												if (key != MsoPictureEffectType.msoEffectColorTemperature || !options.Picture.Temperature)
												{
													continue;
												}
												while (true)
												{
													switch (3)
													{
													case 0:
														continue;
													}
													break;
												}
												shape.Fill.PictureEffects.Insert(MsoPictureEffectType.msoEffectColorTemperature).EffectParameters[1].Value = current.Value[0];
												_ = null;
												continue;
											}
											flag = true;
											if (!options.Picture.Brightness)
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
												if (!options.Picture.Contrast)
												{
													continue;
												}
												while (true)
												{
													switch (7)
													{
													case 0:
														continue;
													}
													break;
												}
											}
											PictureEffect pictureEffect2 = shape.Fill.PictureEffects.Insert(MsoPictureEffectType.msoEffectBrightnessContrast);
											if (options.Picture.Brightness)
											{
												pictureEffect2.EffectParameters[1].Value = current.Value[0];
											}
											if (options.Picture.Contrast)
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
												pictureEffect2.EffectParameters[2].Value = current.Value[1];
											}
											pictureEffect2 = null;
											continue;
										}
										switch (key)
										{
										case MsoPictureEffectType.msoEffectSharpenSoften:
											while (true)
											{
												switch (3)
												{
												case 0:
													continue;
												}
												break;
											}
											if (options.Picture.Sharpness)
											{
												shape.Fill.PictureEffects.Insert(MsoPictureEffectType.msoEffectSharpenSoften).EffectParameters[1].Value = current.Value[0];
												_ = null;
											}
											break;
										case MsoPictureEffectType.msoEffectSaturation:
											if (!options.Picture.Saturation)
											{
												break;
											}
											while (true)
											{
												switch (2)
												{
												case 0:
													continue;
												}
												break;
											}
											shape.Fill.PictureEffects.Insert(MsoPictureEffectType.msoEffectSaturation).EffectParameters[1].Value = current.Value[0];
											_ = null;
											break;
										}
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											break;
										default:
											goto end_IL_0e05;
										}
										continue;
										end_IL_0e05:
										break;
									}
								}
								if (!flag)
								{
									try
									{
										if (options.Picture.Brightness)
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
											pictureFormat.Brightness = picture.Brightness;
										}
										if (options.Picture.Contrast)
										{
											while (true)
											{
												switch (2)
												{
												case 0:
													continue;
												}
												pictureFormat.Contrast = picture.Contrast;
												break;
											}
										}
									}
									catch (Exception ex5)
									{
										ProjectData.SetProjectError(ex5);
										Exception ex6 = ex5;
										ProjectData.ClearProjectError();
									}
								}
								goto IL_0e84;
								IL_0e84:
								if (options.Picture.Crop)
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
									pictureFormat.CropTop = picture.CropTop;
									pictureFormat.CropBottom = picture.CropBottom;
									pictureFormat.CropLeft = picture.CropLeft;
									pictureFormat.CropRight = picture.CropRight;
								}
								if (!options.Picture.ScaleHeight)
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
									if (!options.Picture.ScaleWidth)
									{
										goto IL_0f82;
									}
								}
								MsoTriState lockAspectRatio2 = shape.LockAspectRatio;
								try
								{
									shape.LockAspectRatio = MsoTriState.msoFalse;
									if (options.Picture.ScaleHeight)
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
										shape.ScaleHeight(picture.ScaleHeight, MsoTriState.msoTrue);
									}
									if (options.Picture.ScaleWidth)
									{
										while (true)
										{
											switch (4)
											{
											case 0:
												continue;
											}
											shape.ScaleWidth(picture.ScaleWidth, MsoTriState.msoTrue);
											break;
										}
									}
								}
								catch (Exception ex7)
								{
									ProjectData.SetProjectError(ex7);
									Exception ex8 = ex7;
									ProjectData.ClearProjectError();
								}
								shape.LockAspectRatio = lockAspectRatio2;
								goto IL_0f82;
								IL_0167:
								if (shape.HasTable == MsoTriState.msoFalse)
								{
									if (!options.Line.Color)
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
										if (!options.Line.Style)
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
											if (!options.Line.Weight)
											{
												goto IL_02ae;
											}
											while (true)
											{
												switch (1)
												{
												case 0:
													continue;
												}
												break;
											}
										}
									}
									Properties.LineProperties line = properties.Line;
									if (line.Visible == MsoTriState.msoTrue)
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
										if (options.Line.Color)
										{
											shape.Line.ForeColor.RGB = line.ForeColor;
											shape.Line.BackColor.RGB = line.BackColor;
										}
										if (options.Line.Style)
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
											shape.Line.Style = line.Style;
											shape.Line.DashStyle = line.DashStyle;
										}
										if (options.Line.Weight)
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
											shape.Line.Weight = line.Weight;
										}
										shape.Line.Visible = MsoTriState.msoTrue;
									}
									else
									{
										shape.Line.Visible = MsoTriState.msoFalse;
									}
								}
								goto IL_02ae;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_0fd1;
								}
								continue;
								end_IL_0fd1:
								break;
							}
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
									(enumerator as IDisposable).Dispose();
									break;
								}
							}
						}
						clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)1, AH.A(138737));
					}
					catch (Exception ex9)
					{
						ProjectData.SetProjectError(ex9);
						Exception ex10 = ex9;
						Forms.ErrorMessage(ex10.Message);
						clsReporting.LogException(ex10);
						ProjectData.ClearProjectError();
					}
				}
				catch (Exception ex11)
				{
					ProjectData.SetProjectError(ex11);
					Exception ex12 = ex11;
					A();
					ProjectData.ClearProjectError();
				}
				application = null;
				shapeRange = null;
				return;
			}
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B, Options C)
	{
		Properties.FillProperties fill = B.Fill;
		if (fill.Visible == MsoTriState.msoTrue)
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
			if (fill.Type != MsoFillType.msoFillMixed)
			{
				if (C.Fill.Color)
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
					A.Fill.ForeColor.RGB = fill.ForeColor;
					A.Fill.BackColor.RGB = fill.BackColor;
				}
				if (C.Fill.Type)
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
					switch (fill.Type)
					{
					case MsoFillType.msoFillSolid:
						A.Fill.Solid();
						break;
					case MsoFillType.msoFillPatterned:
						A.Fill.Patterned(fill.Pattern);
						break;
					case MsoFillType.msoFillGradient:
						switch (fill.GradientColorType)
						{
						case MsoGradientColorType.msoGradientTwoColors:
						case MsoGradientColorType.msoGradientMultiColor:
						{
							A.Fill.TwoColorGradient(fill.GradientStyle, fill.GradientVariant);
							A.Fill.GradientAngle = fill.GradientAngle;
							int count = fill.GradientStops.Count;
							for (int i = 1; i <= count; i = checked(i + 1))
							{
								if (i < 3)
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
									GradientStop gradientStop = fill.GradientStops[i];
									if (gradientStop.Color.Type == MsoColorType.msoColorTypeScheme)
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
										A.Fill.GradientStops[i].Color.SchemeColor = gradientStop.Color.SchemeColor;
									}
									else
									{
										A.Fill.GradientStops[i].Color.RGB = gradientStop.Color.RGB;
									}
									A.Fill.GradientStops[i].Position = gradientStop.Position;
									A.Fill.GradientStops[i].Transparency = gradientStop.Transparency;
									A.Fill.GradientStops[i].Color.Brightness = gradientStop.Color.Brightness;
									gradientStop = null;
								}
								else
								{
									GradientStop gradientStop2 = fill.GradientStops[i];
									if (gradientStop2.Color.Type == MsoColorType.msoColorTypeScheme)
									{
										ThemeColorScheme themeColorScheme = NG.A.Application.ActivePresentation.Designs[1].SlideMaster.Theme.ThemeColorScheme;
										A.Fill.GradientStops.Insert2(themeColorScheme.Colors((MsoThemeColorSchemeIndex)gradientStop2.Color.ObjectThemeColor).RGB, gradientStop2.Position, gradientStop2.Transparency, i, gradientStop2.Color.Brightness);
										themeColorScheme = null;
									}
									else
									{
										A.Fill.GradientStops.Insert2(gradientStop2.Color.RGB, gradientStop2.Position, gradientStop2.Transparency, i, gradientStop2.Color.Brightness);
									}
									gradientStop2 = null;
								}
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								break;
							}
							break;
						}
						case MsoGradientColorType.msoGradientOneColor:
							A.Fill.OneColorGradient(fill.GradientStyle, fill.GradientVariant, fill.GradientDegree);
							A.Fill.GradientAngle = fill.GradientAngle;
							break;
						case MsoGradientColorType.msoGradientPresetColors:
							A.Fill.PresetGradient(fill.GradientStyle, fill.GradientVariant, fill.PresetGradientType);
							break;
						}
						break;
					}
				}
				if (C.Fill.Transparency)
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
					try
					{
						A.Fill.Transparency = fill.Transparency;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				A.Fill.Visible = MsoTriState.msoTrue;
				goto IL_0433;
			}
		}
		A.Fill.Visible = MsoTriState.msoFalse;
		goto IL_0433;
		IL_0433:
		fill = default(Properties.FillProperties);
	}

	private static void A(TextRange2 A, Properties B, Options C)
	{
		ParagraphFormat2 paragraphFormat = A.ParagraphFormat;
		Font2 font = A.Font;
		Properties.FontProperties font2 = B.Font;
		string name = font.Name;
		float size = font.Size;
		if (C.Font.Color)
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
			font.UnderlineColor.RGB = font2.UnderlineColor;
			if (font2.Highlight != 0)
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
				if (font.Highlight.RGB != 0)
				{
					font.Highlight.RGB = font2.Highlight;
				}
			}
		}
		if (C.Font.Decoration)
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
			Properties.DecorationProperties decoration = font2.Decoration;
			font.Bold = decoration.Bold;
			font.Italic = decoration.Italic;
			font.UnderlineStyle = decoration.UnderlineStyle;
			font.Strike = decoration.Strike;
			font.StrikeThrough = decoration.StrikeThrough;
			font.DoubleStrikeThrough = decoration.DoubleStrikeThrough;
		}
		if (C.Font.Color)
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
			font.Fill.ForeColor.RGB = font2.ForeColor;
			font.Fill.BackColor.RGB = font2.BackColor;
		}
		if (C.Font.Name)
		{
			font.Name = font2.Name;
		}
		else if (C.Font.Color)
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
			font.Name = name;
		}
		if (C.Font.Size)
		{
			font.Size = font2.Size;
		}
		else if (C.Font.Color)
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
			font.Size = size;
		}
		font2 = default(Properties.FontProperties);
		Properties.TextBoxProperties textBox = B.TextBox;
		if (C.TextBox.Bullets)
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
			ApplyBullets(A, B.TextBox);
		}
		if (C.TextBox.Indents)
		{
			ApplyIndents(A, B.TextBox);
		}
		if (C.TextBox.LineSpacing)
		{
			ApplyLineSpacing(A, B.TextBox);
		}
		if (C.TextBox.HorizontalAlignment)
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
			paragraphFormat.Alignment = textBox.HorizontalAlignment;
		}
		textBox = default(Properties.TextBoxProperties);
		Properties.TextEffectsProperties textEffects = B.TextEffects;
		if (C.TextEffects.Glow)
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
			Apply.A(font.Glow, textEffects.Glow);
		}
		if (C.TextEffects.Reflection)
		{
			Apply.A(font.Reflection, textEffects.Reflection);
		}
		if (C.TextEffects.Shadow)
		{
			Apply.A(font.Shadow, textEffects.Shadow);
		}
		if (C.TextEffects.SoftEdge)
		{
			font.SoftEdgeFormat = textEffects.SoftEdge;
		}
		textEffects = default(Properties.TextEffectsProperties);
		paragraphFormat = null;
		font = null;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.TextFrame2 A, Properties B, Options C)
	{
		Properties.TextBoxProperties textBox = B.TextBox;
		if (C.TextBox.Margins)
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
			A.MarginTop = textBox.MarginTop;
			A.MarginBottom = textBox.MarginBottom;
			A.MarginLeft = textBox.MarginLeft;
			A.MarginRight = textBox.MarginRight;
		}
		if (C.TextBox.AutoSize)
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
			A.AutoSize = textBox.AutoSize;
		}
		if (C.TextBox.WordWrap)
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
			A.WordWrap = textBox.WordWrap;
		}
		if (C.TextBox.HorizontalAlignment)
		{
			A.HorizontalAnchor = textBox.HorizontalAnchor;
		}
		if (C.TextBox.VerticalAlignment)
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
			A.VerticalAnchor = textBox.VerticalAnchor;
		}
		if (C.TextBox.Orientation)
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
			A.Orientation = textBox.Orientation;
		}
		textBox = default(Properties.TextBoxProperties);
		Properties.TextEffectsProperties textEffects = B.TextEffects;
		if (C.TextEffects.ThreeD)
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
			Apply.A(A.ThreeD, textEffects.ThreeD);
		}
		textEffects = default(Properties.TextEffectsProperties);
	}

	public static void ApplyBullets(TextRange2 rng, Properties.TextBoxProperties props)
	{
		Dictionary<int, Properties.BulletProperties> bullets = props.Bullets;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = rng.get_Paragraphs(-1, -1).GetEnumerator();
			while (enumerator.MoveNext())
			{
				ParagraphFormat2 paragraphFormat = ((TextRange2)enumerator.Current).ParagraphFormat;
				if (bullets.TryGetValue(paragraphFormat.IndentLevel, out var value))
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
					paragraphFormat.Alignment = MsoParagraphAlignment.msoAlignLeft;
					Font2 font = value.Font;
					BulletFormat2 bullet = paragraphFormat.Bullet;
					bullet.RelativeSize = value.RelativeSize;
					bullet.Type = value.Type;
					switch (value.Type)
					{
					case MsoBulletType.msoBulletUnnumbered:
						bullet.Character = value.Character;
						break;
					case MsoBulletType.msoBulletNumbered:
						bullet.StartValue = value.StartValue;
						bullet.Style = value.Style;
						break;
					case MsoBulletType.msoBulletMixed:
						bullet.Style = MsoNumberedBulletStyle.msoBulletStyleMixed;
						break;
					}
					Font2 font2 = bullet.Font;
					font2.Name = font.Name;
					font2.Bold = font.Bold;
					font2.Fill.ForeColor = font.Fill.ForeColor;
					font2.Fill.BackColor = font.Fill.BackColor;
					_ = null;
					bullet.UseTextColor = value.UseTextColor;
					if (value.UseTextColor == MsoTriState.msoFalse)
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
						bullet.Font.Fill.ForeColor.RGB = font.Fill.ForeColor.RGB;
					}
					bullet.UseTextFont = value.UseTextFont;
					if (value.UseTextFont == MsoTriState.msoFalse)
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
						bullet.Font.Name = font.Name;
					}
					bullet = null;
					font = null;
				}
				paragraphFormat = null;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					return;
				}
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

	public static void ApplyIndents(TextRange2 rng, Properties.TextBoxProperties props)
	{
		Dictionary<int, Properties.IndentProperties> indents = props.Indents;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = rng.get_Paragraphs(-1, -1).GetEnumerator();
			while (enumerator.MoveNext())
			{
				ParagraphFormat2 paragraphFormat = ((TextRange2)enumerator.Current).ParagraphFormat;
				if (indents.TryGetValue(paragraphFormat.IndentLevel, out var value))
				{
					paragraphFormat.LeftIndent = value.LeftIndent;
					paragraphFormat.RightIndent = value.RightIndent;
					paragraphFormat.FirstLineIndent = value.FirstLineIndent;
				}
				paragraphFormat = null;
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
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (6)
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

	public static void ApplyLineSpacing(TextRange2 rng, Properties.TextBoxProperties props)
	{
		ParagraphFormat2 paragraphFormat = rng.ParagraphFormat;
		ref Properties.SpacingProperties lineSpacing = ref props.LineSpacing;
		paragraphFormat.SpaceAfter = lineSpacing.SpaceAfter;
		paragraphFormat.SpaceBefore = lineSpacing.SpaceBefore;
		paragraphFormat.LineRuleWithin = lineSpacing.LineRuleWithin;
		paragraphFormat.SpaceWithin = lineSpacing.SpaceWithin;
	}

	private static void A()
	{
		Forms.WarningMessage(AH.A(73308));
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.ThreeDFormat A, Properties.ThreeDProperties B)
	{
		try
		{
			A.Visible = B.Visible;
			if (B.Visible != MsoTriState.msoTrue)
			{
				return;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				A.BevelBottomType = B.BevelBottomType;
				if (B.BevelBottomType != MsoBevelType.msoBevelNone)
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
					A.BevelBottomDepth = B.BevelBottomDepth;
					A.BevelBottomInset = B.BevelBottomInset;
				}
				A.BevelTopType = B.BevelTopType;
				if (B.BevelTopType != MsoBevelType.msoBevelNone)
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
					A.BevelTopDepth = B.BevelTopDepth;
					A.BevelTopInset = B.BevelTopInset;
				}
				A.ContourColor.RGB = B.ContourColor;
				A.ContourWidth = B.ContourWidth;
				A.ExtrusionColor.RGB = B.ExtrusionColor;
				A.ExtrusionColorType = B.ExtrusionColorType;
				A.Depth = B.Depth;
				A.FieldOfView = B.FieldOfView;
				A.LightAngle = B.LightAngle;
				A.Perspective = B.Perspective;
				A.PresetLighting = B.PresetLighting;
				A.PresetLightingDirection = B.PresetLightingDirection;
				A.PresetLightingSoftness = B.PresetLightingSoftness;
				A.PresetMaterial = B.PresetMaterial;
				A.RotationX = B.RotationX;
				A.RotationY = B.RotationY;
				A.RotationZ = B.RotationZ;
				A.ProjectText = B.ProjectText;
				A.Z = B.Z;
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.ShadowFormat A, Properties.ShapeShadowProperties B)
	{
		try
		{
			A.Visible = B.Visible;
			if (B.Visible != MsoTriState.msoTrue)
			{
				return;
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
				A.Type = B.Type;
				A.Style = B.Style;
				A.Blur = B.Blur;
				A.ForeColor.RGB = B.ForeColor;
				A.Obscured = B.Obscured;
				A.OffsetX = B.OffsetX;
				A.OffsetY = B.OffsetY;
				A.RotateWithShape = B.RotateWithShape;
				A.Size = B.Size;
				A.Transparency = B.Transparency;
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(Microsoft.Office.Core.ShadowFormat A, Properties.TextShadowProperties B)
	{
		try
		{
			A.Visible = B.Visible;
			if (B.Visible != MsoTriState.msoTrue)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				A.Type = B.Type;
				A.Style = B.Style;
				A.Blur = B.Blur;
				A.ForeColor.RGB = B.ForeColor;
				A.Obscured = B.Obscured;
				A.OffsetX = B.OffsetX;
				A.OffsetY = B.OffsetY;
				A.RotateWithShape = B.RotateWithShape;
				A.Size = B.Size;
				A.Transparency = B.Transparency;
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(ReflectionFormat A, Properties.ReflectionProperties B)
	{
		try
		{
			A.Type = B.Type;
			if (B.Type != MsoReflectionType.msoReflectionTypeNone)
			{
				A.Blur = B.Blur;
				A.Offset = B.Offset;
				A.Size = B.Size;
				A.Transparency = B.Transparency;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(GlowFormat A, Properties.GlowProperties B)
	{
		try
		{
			A.Color.RGB = B.Color;
			A.Radius = B.Radius;
			A.Transparency = B.Transparency;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
