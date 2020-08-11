using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SAPbobsCOM;

namespace SAP.Models
{
    public class B1SalesInvoice
    {
        internal static string create_IV(Company oCompany, B1SalesBlanketAgreement model)
        {
            Documents oIV = (Documents)oCompany.GetBusinessObject(BoObjectTypes.oInvoices);
            Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string FiscalTax = "";
            int SellerCode = 0;

            try
            {
                string sql = string.Format("select    top 1 " +
                                           "          pes.CardCode as CardCode, " +
                                           "          pes.CardName as CardName, " +
                                           "          descob.Name as City, " +
                                           "          estcob.Code as State, " +
                                           "          paicob.Code as Country " +
                                           "from      [dbo].[OCRD] pes " +
                                           "          left join[dbo].[CRD1] adrcob on(pes.CardCode = adrcob.CardCode and adrcob.AdresType = 'S' and adrcob.Address = pes.ShipToDef) " +
                                           "          left join[dbo].[OCNT] descob on(descob.AbsId = adrcob.County) " +
                                           "          left join[dbo].[OCST] estcob on(estcob.Code = adrcob.State and estcob.Country = adrcob.Country) " +
                                           "          left join[dbo].[OCRY] paicob on(paicob.Code = adrcob.Country) " +
                                           "          left join[dbo].[CRD1] adrdes on(pes.CardCode = adrdes.CardCode and adrdes.AdresType = 'B' and adrdes.Address = pes.BillToDef) " +
                                           "          left join[dbo].[OCNT] desdes on(desdes.AbsId = adrdes.County) " +
                                           "          left join[dbo].[OCST] estdes on(estdes.Code = adrdes.State and estdes.Country = adrdes.Country) " +
                                           "          left join[dbo].[OCRY] paides on(paides.Code = adrdes.Country) " +
                                           "where     pes.CardCode = '" + model.CardCode + "' ");

                rs.DoQuery(sql);

                if (rs.Fields.Item("Country").Value != "BR")
                {
                    FiscalTax = "9207"; //Fora de Florianópolis estrangeiro
                }
                else if (rs.Fields.Item("Country").Value == "BR" & rs.Fields.Item("State").Value != "SC")
                {
                    FiscalTax = "9203"; //Fora de SC
                }
                else if (rs.Fields.Item("Country").Value == "BR" & rs.Fields.Item("State").Value == "SC" & rs.Fields.Item("City").Value != "Florianópolis")
                {
                    FiscalTax = "9202"; //Fora de Florianópolis mas de SC
                }
                else if (rs.Fields.Item("Country").Value == "BR" & rs.Fields.Item("State").Value == "SC" & rs.Fields.Item("City").Value == "Florianópolis")
                {
                    FiscalTax = "9201"; //De florianópolis
                }
                else
                {
                    FiscalTax = "9201";
                }
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                throw new Exception($"{errCode}-{errMsg}");
            }

            try
            {
                string sql = string.Format("select    SlpCode                                  " +
                                           "from      [dbo].[OSLP]                             " +
                                           "where     Email like '%" + model.SellerEmail + "%' " +
                                           "and       Active = 'Y'                             ");

                rs.DoQuery(sql);

                SellerCode = rs.Fields.Item("SlpCode").Value;
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                throw new Exception($"{errCode}-{errMsg}");
            }

            try
            {
                BusinessPartners oBP = (BusinessPartners)oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                Recordset rs_payment = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                if (oBP.GetByKey(model.CardCode))
                {
                    string sql = string.Format("select    PayMethCod " +
                                               "from      [dbo].[OPYM] " +
                                               "where     Type = 'I' " +
                                               "and       Active = 'Y' ");

                    rs_payment.DoQuery(sql);
                    for (int i = 0; i < rs_payment.RecordCount; i++)
                    {
                        oBP.BPPaymentMethods.PaymentMethodCode = rs_payment.Fields.Item("PayMethCod").Value;
                        oBP.BPPaymentMethods.Add();
                        rs_payment.MoveNext();
                    }
                    oBP.PeymentMethodCode = "Paypal";
                    if (oBP.Update() != 0)
                    {
                        int errCode;
                        string errMsg;
                        oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception($"{errCode}-{errMsg}");
                    }
                }
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                throw new Exception($"{errCode}-{errMsg}");
            }

            try
            {
                oIV.CardCode = model.CardCode;
                oIV.DocDate = model.InitialDate;
                oIV.DocDueDate = model.InitialDate;
                oIV.TaxDate = model.InitialDate;
                if (model.Currency == "USD" | model.Currency == "EUR")
                {
                    oIV.DocCurrency = model.Currency;
                    oIV.DocRate = model.ValueRate;
                }
                oIV.SequenceCode = 49;
                oIV.GroupNumber = -1;
                oIV.UserFields.Fields.Item("U_SKILL_NrCFPS").Value = FiscalTax;
                oIV.UserFields.Fields.Item("U_SKILL_TipTrib").Value = "0";
                oIV.UserFields.Fields.Item("U_SKILL_ServPais").Value = "0";
                oIV.PaymentMethod = model.Origin;
                //Products
                oIV.Lines.SetCurrentLine(0);
                if (model.Currency == "USD" | model.Currency == "EUR")
                {
                    oIV.Lines.Currency = model.Currency;
                }
                oIV.Lines.ItemCode = model.ItemCode;
                oIV.Lines.Quantity = Convert.ToDouble(1);
                oIV.Lines.UnitPrice = model.GrossAmount;
                
                oIV.Lines.WarehouseCode = "01.001";
                if (SellerCode != 0)
                {
                    oIV.Lines.SalesPersonCode = SellerCode;
                    oIV.SalesPersonCode = SellerCode;
                }
                else
                {
                    oIV.Lines.SalesPersonCode = 211;
                    oIV.SalesPersonCode = 211;
                }
                oIV.Lines.CostingCode = "1.06.09";
                if (model.Origin == "Hotmart")
                {
                    if (FiscalTax == "9207")
                    {
                        oIV.Lines.TaxCode = FiscalTax + "0001";
                    }
                    else
                    {
                        oIV.Lines.TaxCode = FiscalTax + "0003";
                    }                    
                }
                else
                {
                    oIV.Lines.TaxCode = FiscalTax + "0001";
                }
                if (model.Origin == "Hotmart")
                {
                    if (FiscalTax == "9207")
                    {
                        oIV.Lines.CFOPCode = "9207";
                    }
                    else if (FiscalTax == "9203")
                    {
                        oIV.Lines.CFOPCode = "9203";
                    }
                    else if (FiscalTax == "9202")
                    {
                        oIV.Lines.CFOPCode = "9202";
                    }
                    else if (FiscalTax == "9201")
                    {
                        oIV.Lines.CFOPCode = "9201";
                    }
                    else
                    {
                        oIV.Lines.CFOPCode = "9207";
                    }
                }
                else
                {
                    if (FiscalTax == "9207")
                    {
                        oIV.Lines.CFOPCode = "9207";
                    }
                    else if (FiscalTax == "9203")
                    {
                        oIV.Lines.CFOPCode = "9203";
                    }
                    else if (FiscalTax == "9202")
                    {
                        oIV.Lines.CFOPCode = "9202";
                    }
                    else if (FiscalTax == "9201")
                    {
                        oIV.Lines.CFOPCode = "9201";
                    }
                    else
                    {
                        oIV.Lines.CFOPCode = "9207";
                    }
                }
                
                oIV.Lines.Usage = "26";
                if (FiscalTax == "9207")
                {
                    oIV.Lines.CSTforPIS = "08";
                    oIV.Lines.CSTforCOFINS = "08";
                }
                else
                {
                    oIV.Lines.CSTforPIS = "01";
                    oIV.Lines.CSTforCOFINS = "01";
                }
                int BlancketNumber = B1SalesBlanketAgreement.read_BA(oCompany, model);
                if (BlancketNumber != 0)
                {
                    oIV.Lines.AgreementNo = BlancketNumber;
                    oIV.Lines.AgreementRowNumber = Convert.ToInt32(0);
                }
                oIV.Lines.Add();

                if (oIV.Add() != 0)
                {
                    int errCode;
                    string errMsg;
                    oCompany.GetLastError(out errCode, out errMsg);
                    throw new Exception($"{errCode}-{errMsg}");
                }
                return Convert.ToString(BlancketNumber) + "|" + oCompany.GetNewObjectKey();
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