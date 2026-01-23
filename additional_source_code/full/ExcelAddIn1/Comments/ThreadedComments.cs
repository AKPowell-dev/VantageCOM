using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Sheets;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Comments;

public sealed class ThreadedComments
{
	public static void DeleteResolved()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		checked
		{
			if (application.Selection is Range)
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
				Range range = (Range)application.Selection;
				if (!A(application, range))
				{
					range = null;
					application = null;
					return;
				}
				application.ScreenUpdating = false;
				application.EnableEvents = false;
				try
				{
					if (application.ActiveWindow.SelectedSheets.Count > 1)
					{
						IEnumerator enumerator = default(IEnumerator);
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							if (MessageBox.Show(VH.A(142593), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
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
								int num = 0;
								int num2 = 0;
								Worksheet worksheet;
								try
								{
									enumerator = application.ActiveWindow.SelectedSheets.GetEnumerator();
									while (enumerator.MoveNext())
									{
										object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
										if (!(objectValue is Worksheet))
										{
											continue;
										}
										worksheet = (Worksheet)objectValue;
										if (!worksheet.ProtectContents)
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
											try
											{
												object obj = NewLateBinding.LateGet(worksheet, null, VH.A(8668), new object[0], null, null, null);
												for (int i = Conversions.ToInteger(NewLateBinding.LateGet(obj, null, VH.A(52690), new object[0], null, null, null)); i >= 1; i += -1)
												{
													object[] array;
													bool[] array2;
													object instance = NewLateBinding.LateGet(obj, null, VH.A(140662), array = new object[1] { i }, null, null, array2 = new bool[1] { true });
													if (array2[0])
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
														i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
													}
													if (!Conversions.ToBoolean(NewLateBinding.LateGet(instance, null, VH.A(102617), new object[0], null, null, null)))
													{
														continue;
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
													object instance2 = obj;
													string memberName = VH.A(140662);
													object[] obj2 = new object[1] { i };
													array = obj2;
													bool[] obj3 = new bool[1] { true };
													array2 = obj3;
													object instance3 = NewLateBinding.LateGet(instance2, null, memberName, obj2, null, null, obj3);
													if (array2[0])
													{
														i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
													}
													NewLateBinding.LateCall(instance3, null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
													num++;
												}
												while (true)
												{
													switch (2)
													{
													case 0:
														continue;
													}
													obj = null;
													break;
												}
											}
											catch (Exception ex)
											{
												ProjectData.SetProjectError(ex);
												Exception ex2 = ex;
												clsReporting.LogException(ex2);
												ProjectData.ClearProjectError();
											}
										}
										num2++;
									}
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											goto end_IL_02c3;
										}
										continue;
										end_IL_02c3:
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
								worksheet = null;
								range.Select();
								if (num > 0)
								{
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										Forms.SuccessMessage(VH.A(52374) + num + VH.A(142745) + num2 + VH.A(142175));
										break;
									}
								}
								else
								{
									Forms.InfoMessage(VH.A(142790));
								}
								break;
							}
							break;
						}
					}
					else
					{
						ExcelAddIn1.Sheets.Protection.Unprotect(range.Worksheet);
						if (!range.Worksheet.ProtectContents)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								if (MessageBox.Show(VH.A(142883), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
								{
									break;
								}
								try
								{
									object instance4 = NewLateBinding.LateGet(range.Worksheet, null, VH.A(8668), new object[0], null, null, null);
									for (int j = Conversions.ToInteger(NewLateBinding.LateGet(instance4, null, VH.A(52690), new object[0], null, null, null)); j >= 1; j += -1)
									{
										object[] array;
										bool[] array2;
										object obj4 = NewLateBinding.LateGet(instance4, null, VH.A(140662), array = new object[1] { j }, null, null, array2 = new bool[1] { true });
										if (array2[0])
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
											j = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
										}
										object instance5 = obj4;
										if (Conversions.ToBoolean(Conversions.ToBoolean(NewLateBinding.LateGet(instance5, null, VH.A(102617), new object[0], null, null, null)) && application.Intersect(range, (Range)NewLateBinding.LateGet(instance5, null, VH.A(8701), new object[0], null, null, null), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null))
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
											NewLateBinding.LateCall(instance5, null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
										}
										instance5 = null;
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											continue;
										}
										instance4 = null;
										break;
									}
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									Forms.ErrorMessage(ex4.Message);
									clsReporting.LogException(ex4);
									ProjectData.ClearProjectError();
								}
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
				application.EnableEvents = true;
				application.ScreenUpdating = true;
				range = null;
			}
			application = null;
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(143023));
		}
	}

	public static void Resolve()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		checked
		{
			if (application.Selection is Range)
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
				Range range = (Range)application.Selection;
				if (!A(application, range))
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						range = null;
						application = null;
						return;
					}
				}
				application.ScreenUpdating = false;
				application.EnableEvents = false;
				try
				{
					if (application.ActiveWindow.SelectedSheets.Count > 1)
					{
						if (MessageBox.Show(VH.A(143072), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
						{
							int num = 0;
							int num2 = 0;
							IEnumerator enumerator = default(IEnumerator);
							Worksheet worksheet;
							try
							{
								enumerator = application.ActiveWindow.SelectedSheets.GetEnumerator();
								while (enumerator.MoveNext())
								{
									object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
									if (!(objectValue is Worksheet))
									{
										continue;
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
									worksheet = (Worksheet)objectValue;
									if (!worksheet.ProtectContents)
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
											object obj = NewLateBinding.LateGet(worksheet, null, VH.A(8668), new object[0], null, null, null);
											for (int i = Conversions.ToInteger(NewLateBinding.LateGet(obj, null, VH.A(52690), new object[0], null, null, null)); i >= 1; i += -1)
											{
												object[] array;
												bool[] array2;
												object instance = NewLateBinding.LateGet(obj, null, VH.A(140662), array = new object[1] { i }, null, null, array2 = new bool[1] { true });
												if (array2[0])
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
													i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
												}
												if (!Conversions.ToBoolean(Operators.NotObject(NewLateBinding.LateGet(instance, null, VH.A(102617), new object[0], null, null, null))))
												{
													continue;
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
												object instance2 = obj;
												string memberName = VH.A(140662);
												object[] obj2 = new object[1] { i };
												array = obj2;
												bool[] obj3 = new bool[1] { true };
												array2 = obj3;
												object instance3 = NewLateBinding.LateGet(instance2, null, memberName, obj2, null, null, obj3);
												if (array2[0])
												{
													i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
												}
												NewLateBinding.LateSetComplex(instance3, null, VH.A(102617), new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
												num++;
											}
											while (true)
											{
												switch (1)
												{
												case 0:
													continue;
												}
												obj = null;
												break;
											}
										}
										catch (Exception ex)
										{
											ProjectData.SetProjectError(ex);
											Exception ex2 = ex;
											clsReporting.LogException(ex2);
											ProjectData.ClearProjectError();
										}
									}
									num2++;
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										goto end_IL_02c4;
									}
									continue;
									end_IL_02c4:
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
							worksheet = null;
							range.Select();
							if (num > 0)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									Forms.SuccessMessage(VH.A(143204) + num + VH.A(143223) + num2 + VH.A(142175));
									break;
								}
							}
							else
							{
								Forms.InfoMessage(VH.A(143250));
							}
						}
					}
					else
					{
						ExcelAddIn1.Sheets.Protection.Unprotect(range.Worksheet);
						if (!range.Worksheet.ProtectContents && MessageBox.Show(VH.A(143347), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
						{
							try
							{
								int num = 0;
								object instance4 = NewLateBinding.LateGet(range.Worksheet, null, VH.A(8668), new object[0], null, null, null);
								for (int j = Conversions.ToInteger(NewLateBinding.LateGet(instance4, null, VH.A(52690), new object[0], null, null, null)); j >= 1; j += -1)
								{
									object[] array;
									bool[] array2;
									object obj4 = NewLateBinding.LateGet(instance4, null, VH.A(140662), array = new object[1] { j }, null, null, array2 = new bool[1] { true });
									if (array2[0])
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
										j = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
									}
									object instance5 = obj4;
									if (Conversions.ToBoolean(Conversions.ToBoolean(Operators.NotObject(NewLateBinding.LateGet(instance5, null, VH.A(102617), new object[0], null, null, null))) && application.Intersect(range, (Range)NewLateBinding.LateGet(instance5, null, VH.A(8701), new object[0], null, null, null), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null))
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
										NewLateBinding.LateSetComplex(instance5, null, VH.A(102617), new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
										num++;
									}
									instance5 = null;
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									instance4 = null;
									if (num != 0)
									{
										while (true)
										{
											switch (2)
											{
											case 0:
												continue;
											}
											if (num != 1)
											{
												while (true)
												{
													switch (6)
													{
													case 0:
														continue;
													}
													Forms.SuccessMessage(VH.A(143204) + num + VH.A(143606));
													break;
												}
											}
											else
											{
												Forms.SuccessMessage(VH.A(143567));
											}
											break;
										}
									}
									else
									{
										Forms.InfoMessage(VH.A(143474));
									}
									break;
								}
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								Forms.ErrorMessage(ex4.Message);
								clsReporting.LogException(ex4);
								ProjectData.ClearProjectError();
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
				application.EnableEvents = true;
				application.ScreenUpdating = true;
				range = null;
			}
			application = null;
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(143627));
		}
	}

	public static void Reopen()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		checked
		{
			if (application.Selection is Range)
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
				Range range = (Range)application.Selection;
				if (!A(application, range))
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						range = null;
						application = null;
						return;
					}
				}
				application.ScreenUpdating = false;
				application.EnableEvents = false;
				try
				{
					if (application.ActiveWindow.SelectedSheets.Count > 1)
					{
						if (MessageBox.Show(VH.A(143660), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
						{
							IEnumerator enumerator = default(IEnumerator);
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								int num = 0;
								int num2 = 0;
								Worksheet worksheet;
								try
								{
									enumerator = application.ActiveWindow.SelectedSheets.GetEnumerator();
									while (enumerator.MoveNext())
									{
										object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
										if (!(objectValue is Worksheet))
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
										worksheet = (Worksheet)objectValue;
										if (!worksheet.ProtectContents)
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
												object obj = NewLateBinding.LateGet(worksheet, null, VH.A(8668), new object[0], null, null, null);
												for (int i = Conversions.ToInteger(NewLateBinding.LateGet(obj, null, VH.A(52690), new object[0], null, null, null)); i >= 1; i += -1)
												{
													object[] array;
													bool[] array2;
													object instance = NewLateBinding.LateGet(obj, null, VH.A(140662), array = new object[1] { i }, null, null, array2 = new bool[1] { true });
													if (array2[0])
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
														i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
													}
													if (!Conversions.ToBoolean(NewLateBinding.LateGet(instance, null, VH.A(102617), new object[0], null, null, null)))
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
													object instance2 = obj;
													string memberName = VH.A(140662);
													object[] obj2 = new object[1] { i };
													array = obj2;
													bool[] obj3 = new bool[1] { true };
													array2 = obj3;
													object instance3 = NewLateBinding.LateGet(instance2, null, memberName, obj2, null, null, obj3);
													if (array2[0])
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
														i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
													}
													NewLateBinding.LateSetComplex(instance3, null, VH.A(102617), new object[1] { false }, null, null, OptimisticSet: false, RValueBase: true);
													num++;
												}
												while (true)
												{
													switch (6)
													{
													case 0:
														continue;
													}
													obj = null;
													break;
												}
											}
											catch (Exception ex)
											{
												ProjectData.SetProjectError(ex);
												Exception ex2 = ex;
												clsReporting.LogException(ex2);
												ProjectData.ClearProjectError();
											}
										}
										num2++;
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
								worksheet = null;
								range.Select();
								if (num > 0)
								{
									Forms.SuccessMessage(VH.A(143808) + num + VH.A(143223) + num2 + VH.A(142175));
								}
								else
								{
									Forms.InfoMessage(VH.A(142790));
								}
								break;
							}
						}
					}
					else
					{
						ExcelAddIn1.Sheets.Protection.Unprotect(range.Worksheet);
						if (!range.Worksheet.ProtectContents && MessageBox.Show(VH.A(143827), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								try
								{
									int num = 0;
									object instance4 = NewLateBinding.LateGet(range.Worksheet, null, VH.A(8668), new object[0], null, null, null);
									for (int j = Conversions.ToInteger(NewLateBinding.LateGet(instance4, null, VH.A(52690), new object[0], null, null, null)); j >= 1; j += -1)
									{
										object[] array;
										bool[] array2;
										object obj4 = NewLateBinding.LateGet(instance4, null, VH.A(140662), array = new object[1] { j }, null, null, array2 = new bool[1] { true });
										if (array2[0])
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
											j = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
										}
										object instance5 = obj4;
										if (Conversions.ToBoolean(Conversions.ToBoolean(NewLateBinding.LateGet(instance5, null, VH.A(102617), new object[0], null, null, null)) && application.Intersect(range, (Range)NewLateBinding.LateGet(instance5, null, VH.A(8701), new object[0], null, null, null), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null))
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
											NewLateBinding.LateSetComplex(instance5, null, VH.A(102617), new object[1] { false }, null, null, OptimisticSet: false, RValueBase: true);
											num++;
										}
										instance5 = null;
									}
									while (true)
									{
										switch (4)
										{
										case 0:
											continue;
										}
										instance4 = null;
										if (num != 0)
										{
											while (true)
											{
												switch (5)
												{
												case 0:
													continue;
												}
												if (num == 1)
												{
													Forms.SuccessMessage(VH.A(144060));
												}
												else
												{
													Forms.SuccessMessage(VH.A(143808) + num + VH.A(143606));
												}
												break;
											}
										}
										else
										{
											Forms.InfoMessage(VH.A(143971));
										}
										break;
									}
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									Forms.ErrorMessage(ex4.Message);
									clsReporting.LogException(ex4);
									ProjectData.ClearProjectError();
								}
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
				application.EnableEvents = true;
				application.ScreenUpdating = true;
				range = null;
			}
			application = null;
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(144099));
		}
	}

	private static bool A(Microsoft.Office.Interop.Excel.Application A, Range B)
	{
		bool flag = true;
		if (Conversion.Val(A.Version) >= 16.0)
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
			try
			{
				Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(NewLateBinding.LateGet(B.Worksheet, null, VH.A(8668), new object[0], null, null, null), null, VH.A(52690), new object[0], null, null, null), 0, TextCompare: false);
			}
			catch (COMException ex)
			{
				ProjectData.SetProjectError(ex);
				COMException ex2 = ex;
				if (ex2.Message.Contains(VH.A(144130)))
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
					flag = false;
				}
				ProjectData.ClearProjectError();
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		else
		{
			flag = false;
		}
		if (!flag)
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
			Forms.ErrorMessage(VH.A(144199));
		}
		return flag;
	}

	public static void ConvertFromNote()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		if (application.Selection is Range)
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
			Range range = (Range)application.Selection;
			if (!A(application, range))
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					range = null;
					application = null;
					return;
				}
			}
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			Microsoft.Office.Interop.Excel.Sheets selectedSheets;
			try
			{
				selectedSheets = application.ActiveWindow.SelectedSheets;
				if (selectedSheets.Count > 1)
				{
					if (MessageBox.Show(VH.A(144322), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
					{
						IEnumerator enumerator = default(IEnumerator);
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							int B = 0;
							int num = 0;
							Worksheet worksheet;
							try
							{
								enumerator = application.ActiveWindow.SelectedSheets.GetEnumerator();
								while (enumerator.MoveNext())
								{
									object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
									if (!(objectValue is Worksheet))
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
										break;
									}
									worksheet = (Worksheet)objectValue;
									if (!worksheet.ProtectContents)
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
										Range range2 = RangeHelpers.B(worksheet);
										if (range2 != null)
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
											worksheet.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
											A(range2, ref B);
											range2 = null;
										}
									}
									num = checked(num + 1);
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_015e;
									}
									continue;
									end_IL_015e:
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
							worksheet = null;
							selectedSheets.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
							range.Worksheet.Activate();
							if (B > 0)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									Forms.SuccessMessage(VH.A(144472) + B + VH.A(144493) + num + VH.A(142175));
									break;
								}
							}
							else
							{
								Forms.InfoMessage(VH.A(144538));
							}
							break;
						}
					}
				}
				else
				{
					ExcelAddIn1.Sheets.Protection.Unprotect(range.Worksheet);
					if (!range.Worksheet.ProtectContents)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							if (MessageBox.Show(VH.A(144607), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
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
								int B = 0;
								Range range2 = RangeHelpers.F(range);
								if (range2 == null)
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
									A(range2, ref B);
									range2 = null;
									break;
								}
								break;
							}
							break;
						}
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(ex2.Message);
				ProjectData.ClearProjectError();
			}
			application.EnableEvents = true;
			application.ScreenUpdating = true;
			range = null;
			selectedSheets = null;
		}
		application = null;
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(144753));
	}

	private static void A(Range A, ref int B)
	{
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range obj = (Range)enumerator.Current;
					string text = Author.RemoveFromText(obj.Comment);
					obj.Comment.Delete();
					object[] array;
					bool[] array2;
					NewLateBinding.LateCall(obj, null, VH.A(117350), array = new object[1] { text }, null, null, array2 = new bool[1] { true }, IgnoreReturn: true);
					if (array2[0])
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
						text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
					}
					_ = null;
					B++;
				}
				while (true)
				{
					switch (4)
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
		}
	}
}
