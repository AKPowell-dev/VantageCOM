using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using A;
using MacabacusMacros.Config;
using MacabacusMacros.Config.Settings;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.MasterShapes;

public sealed class Stamps
{
	private static readonly string m_A = AH.A(150808);

	private static readonly string m_B = AH.A(150815);

	private static readonly string C = AH.A(150858);

	public static string Menu(IRibbonControl control)
	{
		StringBuilder A = new StringBuilder(AH.A(47526));
		List<string> list = new List<string>();
		int num = 0;
		XmlDocument settingsXml = KG.A.SettingsXml;
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = settingsXml.GetElementsByTagName(AH.A(149769)).GetEnumerator();
				while (enumerator.MoveNext())
				{
					XmlElement xmlElement = (XmlElement)enumerator.Current;
					num++;
					list.Add(xmlElement.InnerText);
					Stamps.A(ref A, xmlElement.InnerText, control.Tag, num);
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
			settingsXml = null;
			string presentationStamp = GetPresentationStamp(NG.A.Application.ActivePresentation);
			if (presentationStamp.Length > 0)
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
				if (!list.Contains(presentationStamp))
				{
					num++;
					Stamps.A(ref A, presentationStamp, control.Tag, num);
				}
				list = null;
			}
			bool flag = SharedSettings.IsSettingEditable(Constants.XML_NEW_PRESENTATION_DEFAULT_STAMP);
			if (!flag)
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
				if (num != 0)
				{
					if (flag && num == 0)
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
						MessageBox.Show(AH.A(150130), AH.A(5874), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					}
					goto IL_01c4;
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
			if (num > 0)
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
				A.Append(AH.A(149780));
			}
			A.Append(AH.A(149855) + control.Tag + AH.A(150011));
			goto IL_01c4;
		}
		IL_01c4:
		A.Append(AH.A(49007));
		return A.ToString();
	}

	private static void A(ref StringBuilder A, string B, string C, int D)
	{
		string text = B.Replace(AH.A(82514), AH.A(82517));
		string text2;
		if (D >= 10)
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
			text2 = D + AH.A(14625) + text;
		}
		else
		{
			text2 = AH.A(82543) + D + AH.A(14625) + text;
		}
		string text3 = text2;
		C = C + Stamps.m_A + B.Replace(AH.A(82514), AH.A(82543));
		A.Append(AH.A(150219) + D + AH.A(47705) + text3 + AH.A(150266) + C + AH.A(82654) + text + AH.A(150377));
	}

	public static void Toggle(IRibbonControl control, bool blnAdd)
	{
		if (!Licensing.AllowMasterShapesOperation())
		{
			return;
		}
		string[] array = Strings.Split(control.Tag, Stamps.m_A);
		if (blnAdd)
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
			A(array[1]);
		}
		else
		{
			A("");
		}
		AddRemove.Toggle(array[0], blnAdd);
	}

	public static bool IsVisible(IRibbonControl control)
	{
		string presentationStamp = GetPresentationStamp(NG.A.Application.ActivePresentation);
		if (presentationStamp.Length == 0)
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
					return false;
				}
			}
		}
		return Operators.CompareString(Strings.Split(control.Tag, Stamps.m_A)[1], presentationStamp, TextCompare: false) == 0;
	}

	public static void Custom(IRibbonControl control)
	{
		string text = "";
		text = GetPresentationStamp(NG.A.Application.ActivePresentation);
		text = Forms.InputBox(AH.A(150517), AH.A(150542), text);
		if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
		{
			text = "";
		}
		if (text.Length > 0)
		{
			A(text);
			AddRemove.Toggle(control.Tag, blnAdd: true);
		}
	}

	private static void A(string A)
	{
		Tags tags = NG.A.Application.ActivePresentation.Tags;
		if (A.Length == 0)
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
			tags.Delete(Stamps.m_B);
		}
		else
		{
			tags.Add(Stamps.m_B, A);
		}
		tags = null;
	}

	public static string GetPresentationStamp(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		string result;
		try
		{
			result = pres.Tags[Stamps.m_B];
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = "";
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool HasStampPlaceholder(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		bool flag = false;
		if (shp.HasTextFrame == MsoTriState.msoTrue)
		{
			int num;
			if (shp.TextFrame2.HasText == MsoTriState.msoTrue)
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
				num = (shp.TextFrame2.TextRange.Text.Contains(Placeholders.PLACEHOLDER_STAMP) ? 1 : 0);
			}
			else
			{
				num = 0;
			}
			flag = (byte)num != 0;
		}
		else if (shp.Type == MsoShapeType.msoGroup)
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
			{
				IEnumerator enumerator = shp.GroupItems.GetEnumerator();
				try
				{
					while (true)
					{
						if (enumerator.MoveNext())
						{
							flag = HasStampPlaceholder((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current);
							if (!flag)
							{
								continue;
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_0096;
								}
								continue;
								end_IL_0096:
								break;
							}
							break;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_00aa;
							}
							continue;
							end_IL_00aa:
							break;
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
		return flag;
	}

	public static void AddToNewPresentation(string strStamp)
	{
		Base.C();
		using Dictionary<string, Microsoft.Office.Interop.PowerPoint.Shape>.Enumerator enumerator = Base.MyMasterShapes.GetEnumerator();
		while (enumerator.MoveNext())
		{
			KeyValuePair<string, Microsoft.Office.Interop.PowerPoint.Shape> current = enumerator.Current;
			if (!HasStampPlaceholder(current.Value))
			{
				continue;
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
				A(strStamp);
				AddRemove.Toggle(current.Key, blnAdd: true);
				return;
			}
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

	public static void HideLegacyStamps()
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Microsoft.Office.Interop.PowerPoint.Shape shape = default(Microsoft.Office.Interop.PowerPoint.Shape);
		IEnumerator enumerator = default(IEnumerator);
		Slide slide = default(Slide);
		IEnumerator enumerator2 = default(IEnumerator);
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
				case 350:
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
							goto IL_003f;
						case 4:
							goto IL_006c;
						case 5:
							goto IL_0079;
						case 6:
							goto IL_0093;
						case 7:
							goto IL_009d;
						case 8:
							goto IL_00a8;
						case 9:
							goto IL_00d2;
						case 10:
							goto IL_00ec;
						case 11:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 12:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_006c:
					num2 = 4;
					if (A(shape))
					{
						goto IL_0079;
					}
					goto IL_009d;
					IL_0007:
					num2 = 2;
					enumerator = NG.A.Application.ActivePresentation.Slides.GetEnumerator();
					goto IL_00d5;
					IL_00d5:
					if (enumerator.MoveNext())
					{
						slide = (Slide)enumerator.Current;
						goto IL_003f;
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
					goto IL_00ec;
					IL_0079:
					num2 = 5;
					shape.TextFrame.TextRange.Text = "";
					goto IL_0093;
					IL_00ec:
					num2 = 10;
					if (!(enumerator is IDisposable))
					{
						break;
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
					(enumerator as IDisposable).Dispose();
					break;
					IL_0093:
					num2 = 6;
					shape.Visible = MsoTriState.msoFalse;
					goto IL_009d;
					IL_00d2:
					num2 = 9;
					goto IL_00d5;
					IL_009d:
					num2 = 7;
					goto IL_009f;
					IL_003f:
					num2 = 3;
					enumerator2 = slide.CustomLayout.Shapes.GetEnumerator();
					goto IL_009f;
					IL_009f:
					if (enumerator2.MoveNext())
					{
						shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
						goto IL_006c;
					}
					goto IL_00a8;
					IL_00a8:
					num2 = 8;
					if (enumerator2 is IDisposable)
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
						(enumerator2 as IDisposable).Dispose();
					}
					goto IL_00d2;
					end_IL_0000_2:
					break;
				}
				num2 = 11;
				A("");
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 350;
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
			switch (6)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	private static bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		return Helpers.IsShapeType(A, AH.A(150565));
	}

	public static void ConvertLegacyStamps()
	{
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
		bool flag = false;
		bool flag2 = false;
		bool flag3 = false;
		bool flag4 = false;
		try
		{
			presentation = NG.A.Application.ActivePresentation;
			flag = Conversions.ToBoolean(presentation.Tags[C]);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (!flag)
		{
			Master slideMaster;
			Microsoft.Office.Interop.PowerPoint.Shape shape;
			List<int> list;
			try
			{
				slideMaster = presentation.Designs[1].SlideMaster;
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = slideMaster.Shapes.GetEnumerator();
					while (true)
					{
						if (enumerator.MoveNext())
						{
							if (B((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current))
							{
								flag = true;
								break;
							}
							continue;
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
							break;
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
				if (!flag)
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
					IEnumerator enumerator2 = default(IEnumerator);
					try
					{
						enumerator2 = slideMaster.CustomLayouts.GetEnumerator();
						IEnumerator enumerator3 = default(IEnumerator);
						while (enumerator2.MoveNext())
						{
							CustomLayout customLayout = (CustomLayout)enumerator2.Current;
							list = new List<int>();
							shape = null;
							for (int i = customLayout.Shapes.Count; i >= 1; i = checked(i + -1))
							{
								try
								{
									if (!A(customLayout.Shapes[i]))
									{
										continue;
									}
									while (true)
									{
										switch (1)
										{
										case 0:
											continue;
										}
										list.Add(Helpers.A(customLayout, customLayout.Shapes[i]));
										break;
									}
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									ProjectData.ClearProjectError();
								}
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
							if (list.Count > 0)
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
								if (!flag3)
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
									if (MessageBox.Show(AH.A(150576), AH.A(5874), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
									{
										while (true)
										{
											switch (6)
											{
											case 0:
												continue;
											}
											flag4 = true;
											break;
										}
										break;
									}
									flag3 = true;
								}
								Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = customLayout.Shapes.Range(list.ToArray());
								if (list.Count == 1)
								{
									shape = shapeRange[1];
								}
								else
								{
									if (!flag2)
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
										try
										{
											enumerator3 = shapeRange.GetEnumerator();
											while (enumerator3.MoveNext())
											{
												((Microsoft.Office.Interop.PowerPoint.Shape)enumerator3.Current).Visible = MsoTriState.msoTrue;
											}
										}
										finally
										{
											if (enumerator3 is IDisposable)
											{
												while (true)
												{
													switch (6)
													{
													case 0:
														continue;
													}
													(enumerator3 as IDisposable).Dispose();
													break;
												}
											}
										}
									}
									shape = shapeRange.Group();
								}
								shapeRange = null;
							}
							if (shape == null)
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
							if (!flag2)
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
								shape.Copy();
								Microsoft.Office.Interop.PowerPoint.Shape shape2 = slideMaster.Shapes.Paste()[1];
								shape2.Visible = MsoTriState.msoFalse;
								if (customLayout.Index == 1)
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
									shape2.Name = AH.A(150738);
								}
								else
								{
									shape2.Name = AH.A(150773);
								}
								shape2.TextFrame2.TextRange.Text = Placeholders.PLACEHOLDER_STAMP;
								shape2.Top = shape.Top;
								shape2.Left = shape.Left;
								shape2 = null;
								flag2 = true;
							}
							shape.Delete();
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (7)
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
				if (!flag4)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						presentation.Tags.Add(C, AH.A(149670));
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
			slideMaster = null;
			shape = null;
			list = null;
		}
		presentation = null;
	}

	private static bool B(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		bool flag = false;
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		if (shape.Type != MsoShapeType.msoPlaceholder)
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
			if (shape.HasTextFrame == MsoTriState.msoTrue)
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
				if (shape.TextFrame2.HasText == MsoTriState.msoTrue && shape.TextFrame2.TextRange.Text.Contains(Placeholders.PLACEHOLDER_STAMP))
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
					flag = true;
				}
			}
			else if (shape.Type == MsoShapeType.msoGroup)
			{
				{
					IEnumerator enumerator = shape.GroupItems.GetEnumerator();
					try
					{
						while (true)
						{
							if (enumerator.MoveNext())
							{
								flag = B((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current);
								if (!flag)
								{
									continue;
								}
								while (true)
								{
									switch (5)
									{
									case 0:
										break;
									default:
										goto end_IL_00b4;
									}
									continue;
									end_IL_00b4:
									break;
								}
								break;
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_00ca;
								}
								continue;
								end_IL_00ca:
								break;
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
		}
		shape = null;
		return flag;
	}
}
