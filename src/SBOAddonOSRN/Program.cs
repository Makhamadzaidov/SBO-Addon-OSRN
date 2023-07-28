using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;

namespace SBOAddonOSRN
{
	class Program
	{
		static SAPbobsCOM.Company oCom;

		[STAThread]
		static void Main(string[] args)
		{
			try
			{
				Application oApp = null;
				if (args.Length < 1)
				{
					oApp = new Application();
				}
				else
				{
					oApp = new Application(args[0]);
				}
				oCom = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

				Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
				Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;

				oApp.Run();
			}
			catch (Exception ex)
			{
				System.Windows.Forms.MessageBox.Show(ex.Message);
			}
		}

		private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
		{
			BubbleEvent = true;

			if (pVal.FormTypeEx == "21" && pVal.BeforeAction == false && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
			{
				SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
				SAPbouiCOM.Item oButtonPurchase = oForm.Items.Add("Click", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
				SAPbouiCOM.Item oTempItem = oForm.Items.Item("2");
				SAPbouiCOM.Button oPostButton = (SAPbouiCOM.Button)oButtonPurchase.Specific;

				oPostButton.Caption = "Load";
				oButtonPurchase.Left = oTempItem.Left + oTempItem.Width + 5;
				oButtonPurchase.Top = oTempItem.Top;
				oButtonPurchase.Width = 100;
				oButtonPurchase.Height = oTempItem.Height;
				oButtonPurchase.AffectsFormMode = false;
			}

			if (pVal.FormTypeEx == "21" && pVal.BeforeAction == false && pVal.ItemUID == "Click" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
			{
				try
				{
					SAPbouiCOM.Form soForm = Application.SBO_Application.Forms.Item(FormUID);
					SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)soForm.Items.Item("43").Specific;
					SAPbouiCOM.Matrix oMatrix2 = (SAPbouiCOM.Matrix)soForm.Items.Item("3").Specific;
					SAPbouiCOM.DBDataSource matrixDT1 = soForm.DataSources.DBDataSources.Item("SBDR");

					SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


					int rowCount = matrixDT1.Size;
					int selectedRow = 0;

					for (int i = 1; i <= rowCount; i++)
					{
						if (oMatrix.IsRowSelected(i))
						{
							selectedRow = i;
						}
					}

					SAPbouiCOM.EditText docQuan = (SAPbouiCOM.EditText)oMatrix.Columns.Item("37").Cells.Item(selectedRow).Specific;
					var itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("5").Cells.Item(selectedRow).Specific).Value.ToString();

					oRS.DoQuery($"SELECT T0.\"ItemCode\", Max(to_number(T0.\"DistNumber\")) FROM OSRN T0 WHERE T0.\"ItemCode\" = '{itemCode}' GROUP BY T0.\"ItemCode\"");

					for (int i = 1; i <= int.Parse(docQuan.Value.ToString().Split('.')[0]); i++)
					{
						if (i > oMatrix2.RowCount)
							oMatrix2.AddRow(1);

						SAPbouiCOM.EditText mnfSerial = (SAPbouiCOM.EditText)oMatrix2.Columns.Item(1).Cells.Item(i).Specific;
						SAPbouiCOM.EditText distNumber = (SAPbouiCOM.EditText)oMatrix2.Columns.Item(2).Cells.Item(i).Specific;
						SAPbouiCOM.EditText date = (SAPbouiCOM.EditText)oMatrix2.Columns.Item("50").Cells.Item(i).Specific;

						if (oRS.RecordCount > 0 && !oRS.EoF)
						{
							var dis = oRS.Fields.Item(1).Value.ToString();

							if (string.IsNullOrEmpty(dis))
							{
								distNumber.Value = $"{i}";
								mnfSerial.Value = itemCode + $"_{i}";
							}
							else
							{
								var number = (long.Parse(oRS.Fields.Item(1).Value.ToString()) + i).ToString();
								Console.WriteLine(number);

								distNumber.Value = (long.Parse(oRS.Fields.Item(1).Value.ToString()) + i).ToString();
								mnfSerial.Value = itemCode + $"_{distNumber.Value}";
							}
						}
						else
						{
							distNumber.Value = $"{i}";
							mnfSerial.Value = itemCode + $"_{i}";
						}


						date.Value = DateTime.Now.ToString("yyyyMMdd");
					}
				}
				catch (Exception)
				{
					Application.SBO_Application.SetStatusBarMessage($"{oCom.GetLastErrorDescription()}", SAPbouiCOM.BoMessageTime.bmt_Medium);
				}
			}
		}

		static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
		{
			switch (EventType)
			{
				case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
					//Exit Add-On
					System.Windows.Forms.Application.Exit();
					break;
				case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
					break;
				case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
					break;
				case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
					break;
				case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
					break;
				default:
					break;
			}
		}
	}
}
