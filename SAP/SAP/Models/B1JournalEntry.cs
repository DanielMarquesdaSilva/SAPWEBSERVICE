using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SAPbobsCOM;

namespace SAP.Models
{
    public class B1JournalEntry
    {
        internal static string create_JE(Company oCompany, B1SalesBlanketAgreement model)
        {
            JournalEntries oJE = (JournalEntries)oCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);
            Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            double CheckValueRate = 0.00;

            try
            {
                if (model.Currency == "USD" | model.Currency == "EUR")
                {
                    try
                    {
                        string sql = string.Format("select    top 1 " +
                                                   "          cur.Rate  as ValueRate " +
                                                   "from      [dbo].[ORTT] cur " +
                                                   "where     convert(varchar, cur.RateDate, 103) = convert(varchar, '" + model.InitialDate.ToString("dd/MM/yyyy") + "', 103) " +
                                                   "and       cur.Currency = '" + model.Currency + "' ");

                        rs.DoQuery(sql);
                        CheckValueRate = rs.Fields.Item("ValueRate").Value;

                        SBObob oSBObob;
                        oSBObob = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                        oSBObob.SetCurrencyRate(model.Currency, model.InitialDate, model.ValueRate, true);
                    }
                    catch (Exception)
                    {
                        int errCode;
                        string errMsg;
                        oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception($"{errCode}-{errMsg}");
                    }
                }

                oJE.TaxDate = model.InitialDate;
                oJE.ReferenceDate = model.InitialDate;
                oJE.DueDate = model.InitialDate;
                oJE.Reference = model.AssignSubscriptionId;
                oJE.Reference2 = model.AssignSubscriptionTransactionId;
                oJE.Memo = model.Origin;
                oJE.Lines.SetCurrentLine(0);
                oJE.Lines.AccountCode = "1.01.01.02.26";
                if (model.Currency == "USD" | model.Currency == "EUR")
                {
                    oJE.Lines.FCDebit = model.NetAmount;
                    oJE.Lines.FCCurrency = model.Currency;
                }
                else
                {
                    oJE.Lines.Debit = model.NetAmount;
                }
                oJE.Lines.Add();
                oJE.Lines.SetCurrentLine(1);
                oJE.Lines.AccountCode = "4.01.01.06.06";
                oJE.Lines.CostingCode = "1.10.04";
                if (model.Currency == "USD" | model.Currency == "EUR")
                {
                    oJE.Lines.FCDebit = model.FeeAmount;
                    oJE.Lines.FCCurrency = model.Currency;
                }
                else
                {
                    oJE.Lines.Debit = model.FeeAmount;
                }
                oJE.Lines.Add();
                oJE.Lines.SetCurrentLine(2);
                oJE.Lines.AccountCode = "1.01.03.01.01";
                oJE.Lines.ShortName = model.CardCode;
                if (model.Currency == "USD" | model.Currency == "EUR")
                {
                    oJE.Lines.FCCredit = model.GrossAmount;
                    oJE.Lines.FCCurrency = model.Currency;
                }
                else
                {
                    oJE.Lines.Credit = model.GrossAmount;
                }
                oJE.Lines.Add();

                if (oJE.Add() != 0)
                {
                    int errCode;
                    string errMsg;
                    oCompany.GetLastError(out errCode, out errMsg);
                    throw new Exception($"{errCode}-{errMsg}");
                }

                if (model.Currency == "USD" | model.Currency == "EUR")
                {
                    SBObob oSBObob;
                    oSBObob = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                    oSBObob.SetCurrencyRate(model.Currency, model.InitialDate, CheckValueRate, true);
                }
                return oCompany.GetNewObjectKey();
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                throw new Exception($"{errCode}-{errMsg}");
            }
        }
    }
}