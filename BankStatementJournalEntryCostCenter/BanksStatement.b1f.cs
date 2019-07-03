using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using Application = SAPbouiCOM.Framework.Application;

namespace BankStatementJournalEntryCostCenter
{
    [FormAttribute("10000005", "BanksStatement.b1f")]
    class BanksStatement : SystemFormBase
    {
        public BanksStatement()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("10000039").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_1000").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.DataAddAfter += new SAPbouiCOM.Framework.FormBase.DataAddAfterHandler(this.Form_DataAddAfter);
            this.ActivateAfter += new SAPbouiCOM.Framework.FormBase.ActivateAfterHandler(this.Form_ActivateAfter);
            this.ClickAfter += new ClickAfterHandler(this.Form_ClickAfter);

        }

        private SAPbouiCOM.Button Button0;

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.ActionSuccess)
            {

            }

        }

        private void OnCustomInitialize()
        {

        }

        private void Form_DataAddAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            if (pVal.ActionSuccess)
            {

            }
        }

        private SAPbouiCOM.Button Button1;

        private void Button1_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            int successCount = 0;
            int hasNoEmployeeCount = 0;
            int hasNoJounralEntryCount = 0;
            int hasNo3130Count = 0;
            int errorCount = 0;
            int totalCount = 0;
            Form activeForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;

            var matrix = (Matrix)activeForm.Items.Item("10000036").Specific;
            activeForm.Freeze(true);

            for (int i = 1; i <= matrix.RowCount; i++)
            {
                if (((ComboBox)matrix.GetCellSpecific("10000037", i)).Selected == null)
                {
                    continue;
                }

                var externalCodeCell = ((ComboBox)matrix.GetCellSpecific("10000037", i)).Selected.Value;
                var internalCodeCell = ((ComboBox)matrix.GetCellSpecific("10000035", i)).Selected.Value;

                Recordset recSet =(Recordset)UIConnection.xCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                recSet.DoQuery($"SELECT InOpCode FROM OBTC WHERE AbsEntry = N'{internalCodeCell}'");
                var internalCode = recSet.Fields.Item("InOpCode").Value.ToString();

                if (externalCodeCell == "PRL1" || internalCode == "PMD - თანამშრ." || internalCode == "PBS - 1430")
                {
                    totalCount++;
                    int journalEntryTransId;
                    try
                    {
                        journalEntryTransId = int.Parse(activeForm.DataSources.DBDataSources.Item(0).GetValue("JDTID", i-1), CultureInfo.InvariantCulture);
                    }
                    catch (Exception e)
                    {
                        hasNoJounralEntryCount++;
                        continue;
                    }
                    EmployeesInfo employee = null;
                    var federalTaxIdCell = ((EditText)matrix.GetCellSpecific("1980000087", i)).Value;
                    try
                    {
                        employee = EmployeeController.Get(federalTaxIdCell);
                    }
                    catch (Exception e)
                    {
                        hasNoEmployeeCount++;
                        continue;
                    }
                    var costCenterCode = employee.CostCenterCode;
                    JournalEntries journalEntry = (SAPbobsCOM.JournalEntries)UIConnection.xCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    journalEntry.GetByKey(journalEntryTransId);

                    bool has3130 = false;

                    for (int j = 0; j < journalEntry.Lines.Count; j++)
                    {
                        journalEntry.Lines.SetCurrentLine(j);
                        if (journalEntry.Lines.AccountCode == "3130" || journalEntry.Lines.AccountCode == "1430")
                        {
                            journalEntry.Lines.CostingCode2 = costCenterCode;
                            has3130 = true;
                        }
                    }

                    var res = journalEntry.Update();

                    if (!has3130)
                    {
                        hasNo3130Count++;
                    }
                    else
                    {
                        if (res == 0)
                        {
                            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetSystemMessage(
                                $"წარმატებით განახლდა საჟურნალო გატარება: {journalEntry.Number}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                            successCount++;
                        }
                        else
                        {
                            Application.SBO_Application.SetStatusBarMessage("საჟურნალო გატარება ვერ განახლდა :" + UIConnection.xCompany.GetLastErrorDescription(),
                                BoMessageTime.bmt_Short, true);
                            errorCount++;
                        }
                    }

                }

            }

            activeForm.Freeze(false);
            Application.SBO_Application.MessageBox(
                $"წარმატებული : {successCount}  {Environment.NewLine} თანამშრომელი არ ყავს : {hasNoEmployeeCount} {Environment.NewLine} 3130 ან 1430 არ აქვს : {hasNo3130Count} {Environment.NewLine} არ აქვს საჟურნალო გატარება : {hasNoJounralEntryCount} {Environment.NewLine} წარუმატებელი : {errorCount} {Environment.NewLine} სულ : {totalCount}");
        }

        private void Form_ActivateAfter(SBOItemEventArg pVal)
        {
            Form activeForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            var status = activeForm.DataSources.DBDataSources.Item("OBNH").GetValue("status", 0);
            GetItem("Item_1000").Enabled = status == "E";
        }

        private void Form_ClickAfter(SBOItemEventArg pVal)
        {
            Form activeForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            var status = activeForm.DataSources.DBDataSources.Item("OBNH").GetValue("status", 0);
            GetItem("Item_1000").Enabled = status == "E";
        }
    }
}
