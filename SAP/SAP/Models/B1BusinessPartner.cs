using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SAPbobsCOM;

namespace SAP.Models
{
    public class B1BusinessPartner
    {
        public string CardCode { get; set; }
        public string CardName { get; set; }
        public string Email { get; set; }
        public string Currency { get; set; }
        public string CNPJ { get; set; }
        public string CPF { get; set; }
        public string TaxId { get; set; }
        public string AddressStreet { get; set; }
        public string AddressNumber { get; set; }
        public string AddressComplement { get; set; }
        public string ZipCode { get; set; }
        public string AddressBlock { get; set; }
        public string CityName { get; set; }
        public string State { get; set; }
        public string Country { get; set; }
        public string PhoneNumber { get; set; }

        internal static string create_BP(Company oCompany, B1BusinessPartner model)
        {
            BusinessPartners oBP = (BusinessPartners)oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);
            string BPState = null;
            int BPCounty = 0;
            //string BP_U_GTTipoPN = null;
            //string BP_U_GToptanteSN = null;
            //string BP_U_SKILL_indIEDest = null;
            //string BP_U_GTorigem = null;
            string BP_DebitorAccount = null;
            string BP_DownPaymentClearAct = null;
            string BP_DownPaymentInterimAccount = null;
            //string BP_CustomerBillofExchangPres = null;
            //string BP_CustomerBillofExchangDisc = null;
            try
            {
                if (model.Country != "BR")
                {
                    //BP_U_GTTipoPN = "5";
                    //BP_U_GToptanteSN = "2";
                    //BP_U_SKILL_indIEDest = "9";
                    //BP_U_GTorigem = "2";
                    BP_DebitorAccount = "1.01.03.01.02";
                    BP_DownPaymentClearAct = "2.01.01.02.02";
                    BP_DownPaymentInterimAccount = "2.01.01.02.02";
                    //BP_CustomerBillofExchangPres = "1.01.03.08.99";
                    //BP_CustomerBillofExchangDisc = "1.01.03.08.99";

                    Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                    try
                    {
                        string sql = string.Format("select    top 1 " +
                                                   "          country.Code as Country_Code, " +
                                                   "          country.Name as Country_Name, " +
                                                   "          state.Code as State_Code, " +
                                                   "          state.Name as State_Name, " +
                                                   "          city.Code as County_Code, " +
                                                   "          city.Name as County_Name, " +
                                                   "          city.AbsId as County_AbsId " +
                                                   "from      [dbo].[OCRY] country " +
                                                   "          left join[dbo].[OCST] state on(country.Code = state.Country) " +
                                                   "          left join[dbo].[OCNT] city on(state.Code = city.State and state.Country = city.Country) " +
                                                   "where     country.Code = '" + model.Country + "' ");
                        rs.DoQuery(sql);
                        BPState = rs.Fields.Item("State_Code").Value;
                        BPCounty = rs.Fields.Item("County_AbsId").Value;
                    }
                    catch (Exception)
                    {
                        int errCode;
                        string errMsg;
                        oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception($"{errCode}-{errMsg}");
                    }
                }
                else if (model.Country == "BR")
                {
                    BP_DebitorAccount = "1.01.03.01.01";
                    BP_DownPaymentClearAct = "2.01.01.02.01";
                    BP_DownPaymentInterimAccount = "2.01.01.02.01";
                    //BP_CustomerBillofExchangPres = "1.01.03.08.99";
                    //BP_CustomerBillofExchangDisc = "1.01.03.08.99";

                    Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                    try
                    {
                        string sql = string.Format("declare @state varchar(30), " +
                                                   "        @county varchar(50); " +
                                                   "declare @businessaddress table(Country_Code nvarchar(50), State_Code nvarchar(50), County_Code int) " +
                                                   "set @state = '" + model.State + "' " +
                                                   "set @county = '" + model.CityName + "' " +
                                                   "insert @businessaddress " +
                                                   "select top 1 " +
                                                   "       cou.Code, " +
                                                   "       sta.Code, " +
                                                   "       cit.AbsId " +
                                                   "from   [dbo].[OCRY] cou " +
                                                   "       inner join[dbo].[OCST] sta on(sta.Country = cou.Code) " +
                                                   "       inner join[dbo].[OCNT] cit on(cit.Country = sta.Country and cit.State = sta.Code) " +
                                                   "where  cou.Code = 'BR' " +
                                                   "and    upper(sta.Code) = upper(@state) " +
                                                   "and    upper(cit.Name) = upper(@county) " +
                                                   "if      " +
                                                   "(select count(*) from @businessaddress) > 0  " +
                                                   "select top 1 " +
                                                   "       County_Code as County_AbsId, " +
                                                   "       State_Code as State_Code " +
                                                   "from   @businessaddress  " +
                                                   "else  " +
                                                   "select top 1  " +
                                                   "       cit.AbsId as County_AbsId, " +
                                                   "       sta.Code as State_Code " +
                                                   "from   [dbo].[OCRY] cou " +
                                                   "       inner join[dbo].[OCST] sta on(sta.Country = cou.Code) " +
                                                   "       inner join[dbo].[OCNT] cit on(cit.Country = sta.Country and cit.State = sta.Code) " +
                                                   "where  cou.Code = 'BR' " +
                                                   "and    upper(sta.Code) = upper(@state) " +
                                                   "order by  1 ");
                        rs.DoQuery(sql);
                        BPState = rs.Fields.Item("State_Code").Value;
                        BPCounty = rs.Fields.Item("County_AbsId").Value;
                    }
                    catch (Exception)
                    {
                        int errCode;
                        string errMsg;
                        oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception($"{errCode}-{errMsg}");
                    }

                //    if (model.State == "SC")
                //    {
                //        if (model.CNPJ != null)
                //        {
                //            BP_U_GTTipoPN = "2";
                //            BP_U_GToptanteSN = "2";
                //            BP_U_SKILL_indIEDest = "1";
                //            BP_U_GTorigem = "1";
                //        }
                //        else
                //        {
                //            BP_U_GTTipoPN = "4";
                //            BP_U_GToptanteSN = "3";
                //            BP_U_SKILL_indIEDest = "9";
                //            BP_U_GTorigem = "1";
                //        }
                //    }
                //    else if (model.State == "MG" | model.State == "PR" | model.State == "RS" | model.State == "RJ" | model.State == "SP")
                //    {
                //        if (model.CNPJ != null)
                //        {
                //            BP_U_GTTipoPN = "1";
                //            BP_U_GToptanteSN = "2";
                //            BP_U_SKILL_indIEDest = "1";
                //            BP_U_GTorigem = "2";
                //        }
                //        else
                //        {
                //            BP_U_GTTipoPN = "3";
                //            BP_U_GToptanteSN = "3";
                //            BP_U_SKILL_indIEDest = "9";
                //            BP_U_GTorigem = "2";
                //        }
                //    }
                //    else
                //    {
                //        if (model.CNPJ != null)
                //        {
                //            BP_U_GTTipoPN = "1";
                //            BP_U_GToptanteSN = "2";
                //            BP_U_SKILL_indIEDest = "1";
                //            BP_U_GTorigem = "2";
                //        }
                //        else
                //        {
                //            BP_U_GTTipoPN = "3";
                //            BP_U_GToptanteSN = "3";
                //            BP_U_SKILL_indIEDest = "9";
                //            BP_U_GTorigem = "2";
                //        }
                //    }
                //}
                //else
                //{
                //    BP_DebitorAccount = "1.01.03.01.01";
                //    BP_DownPaymentClearAct = "2.01.01.02.01";
                //    BP_DownPaymentInterimAccount = "2.01.01.02.01";
                //    BP_CustomerBillofExchangPres = "1.01.03.08.99";
                //    BP_CustomerBillofExchangDisc = "1.01.03.08.99";
                //    BPState = "SC";
                //    BPCounty = 4499;
                //    model.CityName = "Florianópolis";
                //    model.Country = "BR";
                //    BP_U_GTTipoPN = "3";
                //    BP_U_GToptanteSN = "2";
                //    BP_U_SKILL_indIEDest = "9";
                //    BP_U_GTorigem = "1";
                }
                oBP.CardName = model.CardName;
                oBP.CardForeignName = model.CardName;
                oBP.GroupCode = 100;
                oBP.EmailAddress = model.Email;
                oBP.Phone1 = model.PhoneNumber;
                oBP.CardType = BoCardTypes.cCustomer;
                oBP.SubjectToWithholdingTax = BoYesNoEnum.tNO;
                oBP.CompanyRegistrationNumber = "1";
                oBP.DebitorAccount = BP_DebitorAccount;
                oBP.DownPaymentClearAct = BP_DownPaymentClearAct;
                oBP.DownPaymentInterimAccount = BP_DownPaymentInterimAccount;
                //oBP.CustomerBillofExchangPres = BP_CustomerBillofExchangPres;
                //oBP.CustomerBillofExchangDisc = BP_CustomerBillofExchangDisc;
                oBP.SalesPersonCode = 1;
                oBP.Currency = "##";
                //oBP.UserFields.Fields.Item("U_GTTipoPN").Value = BP_U_GTTipoPN;
                //oBP.UserFields.Fields.Item("U_GToptanteSN").Value = BP_U_GToptanteSN;
                //oBP.UserFields.Fields.Item("U_SKILL_indIEDest").Value = BP_U_SKILL_indIEDest;
                //oBP.UserFields.Fields.Item("U_GTorigem").Value = BP_U_GTorigem;
                //oBP.UserFields.Fields.Item("U_AD_StatusBitrix").Value = "Novo";
                oBP.Series = 56;
                oBP.Territory = 1;

                //Bill Address
                oBP.Addresses.AddressType = BoAddressType.bo_BillTo;
                oBP.Addresses.AddressName = "BILL";
                oBP.Addresses.TypeOfAddress = "Rua";
                oBP.Addresses.Street = model.AddressStreet;
                oBP.Addresses.StreetNo = model.AddressNumber;
                oBP.Addresses.BuildingFloorRoom = model.AddressComplement;
                oBP.Addresses.ZipCode = model.ZipCode;
                oBP.Addresses.Block = model.AddressBlock;
                oBP.Addresses.City = model.CityName;
                oBP.Addresses.State = BPState;
                oBP.Addresses.County = Convert.ToString(BPCounty);
                oBP.Addresses.Country = model.Country;
                //oBP.Addresses.UserFields.Fields.Item("U_SKILL_indIEDest").Value = BP_U_SKILL_indIEDest;
                oBP.Addresses.Add();

                //Ship Address
                oBP.Addresses.AddressType = BoAddressType.bo_ShipTo;
                oBP.Addresses.AddressName = "SHIP";
                oBP.Addresses.TypeOfAddress = "Rua";
                oBP.Addresses.Street = model.AddressStreet;
                oBP.Addresses.StreetNo = model.AddressNumber;
                oBP.Addresses.BuildingFloorRoom = model.AddressComplement;
                oBP.Addresses.ZipCode = model.ZipCode;
                oBP.Addresses.Block = model.AddressBlock;
                oBP.Addresses.City = model.CityName;
                oBP.Addresses.State = BPState;
                oBP.Addresses.County = Convert.ToString(BPCounty);
                oBP.Addresses.Country = model.Country;
                //oBP.Addresses.UserFields.Fields.Item("U_SKILL_indIEDest").Value = BP_U_SKILL_indIEDest;
                oBP.Addresses.Add();

                //Fiscal
                oBP.FiscalTaxID.SetCurrentLine(0);
                if (model.CNPJ != null)
                {
                    oBP.FiscalTaxID.TaxId0 = model.CNPJ;
                    oBP.FiscalTaxID.TaxId1 = "Isento";
                    oBP.FiscalTaxID.CNAECode = 255;
                }
                oBP.FiscalTaxID.TaxId4 = model.CPF;
                oBP.FiscalTaxID.TaxId5 = model.TaxId;
                oBP.FiscalTaxID.Add();

                //Payment Method
                Recordset rs_payment = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                try
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
                }
                catch (Exception)
                {
                    int errCode;
                    string errMsg;
                    oCompany.GetLastError(out errCode, out errMsg);
                    throw new Exception($"{errCode}-{errMsg}");
                }
                oBP.PeymentMethodCode = "Boleto-Itau";

                //Contact Person
                oBP.ContactEmployees.SetCurrentLine(0);
                oBP.ContactEmployees.Name = "Integration";
                oBP.ContactEmployees.FirstName = model.CardName.Substring(0, model.CardName.IndexOf(" ", 0));
                oBP.ContactEmployees.E_Mail = model.Email;
                oBP.ContactEmployees.Active = BoYesNoEnum.tYES;
                oBP.ContactEmployees.Address = model.AddressStreet;
                oBP.ContactEmployees.Phone1 = model.PhoneNumber;
                oBP.ContactEmployees.Add();
                
                //Insert
                if (oBP.Add() != 0)
                {
                    int errCode;
                    string errMsg;
                    oCompany.GetLastError(out errCode, out errMsg);
                    throw new Exception($"{errCode}-{errMsg}");
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

        internal static string read_BP(Company oCompany, B1BusinessPartner model)
        {
            BusinessPartners oBP = (BusinessPartners)oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);
            Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                string sql = string.Format("declare @CardCode varchar(30), " +
                                           "        @CNPJ varchar(50), " +
                                           "        @CPF varchar(50), " +
                                           "        @TaxId varchar(50) " +
                                           "declare @businesspartner table(CardCode nvarchar(50)) " +
                                           "set @CardCode = '" + model.CardCode + "' " +
                                           "set @CNPJ = '" + model.CNPJ + "' " +
                                           "set @CPF = '" + model.CPF + "' " +
                                           "set @TaxId = '" + model.TaxId + "' " +
                                           "if len(@CardCode) > 0 insert @businesspartner select top 1 " +
                                           "                                                     pes.CardCode " +
                                           "                                              from   [dbo].[OCRD] pes " +
                                           "                                              where  pes.CardCode = @CardCode " +
                                           "                                              and    pes.CardType in ('C','L') " +
                                           "else if len(@CNPJ) > 0 insert @businesspartner select top 1 " +
                                           "                                                      fis.CardCode " +
                                           "                                               from   [dbo].[CRD7] fis " +
                                           "                                                      inner join [dbo].[OCRD] pes on (pes.CardCode = fis.CardCode)  " +
                                           "                                               where  replace(replace(replace(fis.TaxId0,'.',''),'/',''),'-','') = @CNPJ " +
                                           "                                               and    pes.CardType in ('C','L') " +
                                           "else if len(@CPF) > 0 insert @businesspartner select top 1 " +
                                           "                                                     fis.CardCode " +
                                           "                                              from   [dbo].[CRD7] fis " +
                                           "                                                     inner join [dbo].[OCRD] pes on (pes.CardCode = fis.CardCode)  " +
                                           "                                              where  replace(replace(replace(fis.TaxId4,'.',''),'/',''),'-','') = @CPF " +
                                           "                                              and    pes.CardType in ('C', 'L') " +
                                           "else if len(@TaxId) > 0 insert @businesspartner select top 1 " +
                                           "                                                       fis.CardCode " +
                                           "                                                from   [dbo].[CRD7] fis " +
                                           "                                                       inner join [dbo].[OCRD] pes on (pes.CardCode = fis.CardCode)  " +
                                           "                                                where  fis.TaxId5 = @TaxId " +
                                           "                                                and    pes.CardType in ('C','L') " + 
                                           "select top 1 CardCode from @businesspartner " );

                rs.DoQuery(sql);
                oBP.GetByKey(rs.Fields.Item("CardCode").Value);
                return oBP.CardCode;
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                throw new Exception($"{errCode}-{errMsg}");
            }
        }
        internal static string update_BP(Company oCompany, B1BusinessPartner model)
        {
            BusinessPartners oBP = (BusinessPartners)oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);
            Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string BPState = null;
            int BPCounty = 0;
            string BP_U_GTTipoPN = null;
            string BP_U_GToptanteSN = null;
            string BP_U_SKILL_indIEDest = null;
            string BP_U_GTorigem = null;
            int LineBilltoDef = 99999;
            int LineShiptoDef = 99999;

            try
            {
                if (oBP.GetByKey(model.CardCode))
                {
                    try
                    {
                        string sql = string.Format("declare @cardcode varchar(30); " +
                                                   "declare @businessadress table(VisOrder int, CardCode nvarchar(50), Address nvarchar(50), AdresType nvarchar(2)) " +
                                                   "set @cardcode = '" + model.CardCode + "'; " +
                                                   "insert @businessadress " +
                                                   "select    ROW_NUMBER() OVER(ORDER BY adr.LineNum ASC) - 1 as VisOrder, " +
                                                   "          adr.CardCode as CardCode, " +
                                                   "          adr.Address as Address, " +
                                                   "          adr.AdresType as AdresType " +
                                                   "from      [dbo].[CRD1] adr " +
                                                   "where     adr.CardCode = @CardCode " +
                                                   "select    pes.CardCode, " +
                                                   "          pes.BillToDef, " +
                                                   "		  case when pes.BillToDef is null then 99999 " +
                                                   "               else (select VisOrder from @businessadress tabadr where pes.CardCode = tabadr.CardCode and tabadr.Address = pes.BillToDef and tabadr.AdresType = 'B') " +
                                                   "          end as VisOrderBill, " +
                                                   "          pes.ShipToDef, " +
                                                   "		  case when pes.ShipToDef is null then 99999 " +
                                                   "               else (select VisOrder from @businessadress tabadr where pes.CardCode = tabadr.CardCode and tabadr.Address = pes.ShipToDef and tabadr.AdresType = 'S') " +
                                                   "          end as VisOrderShip " +
                                                   "from      [dbo].[OCRD] pes " +
                                                   "where     pes.CardCode = @cardcode ");

                        rs.DoQuery(sql);
                        LineBilltoDef = rs.Fields.Item("VisOrderBill").Value;
                        LineShiptoDef = rs.Fields.Item("VisOrderShip").Value;
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
                        string sql = string.Format("declare @state varchar(30), " +
                                                   "        @county varchar(50); " +
                                                   "declare @businessaddress table(Country_Code nvarchar(50), State_Code nvarchar(50), County_Code int) " +
                                                   "set @state = '" + model.State + "' " +
                                                   "set @county = '" + model.CityName + "' " +
                                                   "insert @businessaddress " +
                                                   "select top 1 " +
                                                   "       cou.Code, " +
                                                   "       sta.Code, " +
                                                   "       cit.AbsId " +
                                                   "from   [dbo].[OCRY] cou " +
                                                   "       inner join[dbo].[OCST] sta on(sta.Country = cou.Code) " +
                                                   "       inner join[dbo].[OCNT] cit on(cit.Country = sta.Country and cit.State = sta.Code) " +
                                                   "where  cou.Code = 'BR' " +
                                                   "and    upper(sta.Code) = upper(@state) " +
                                                   "and    upper(cit.Name) = upper(@county) " +
                                                   "if      " +
                                                   "(select count(*) from @businessaddress) > 0  " +
                                                   "select top 1 " +
                                                   "       County_Code as County_AbsId, " +
                                                   "       State_Code as State_Code " +
                                                   "from   @businessaddress  " +
                                                   "else  " +
                                                   "select top 1  " +
                                                   "       cit.AbsId as County_AbsId, " +
                                                   "       sta.Code as State_Code " +
                                                   "from   [dbo].[OCRY] cou " +
                                                   "       inner join[dbo].[OCST] sta on(sta.Country = cou.Code) " +
                                                   "       inner join[dbo].[OCNT] cit on(cit.Country = sta.Country and cit.State = sta.Code) " +
                                                   "where  cou.Code = 'BR' " +
                                                   "and    upper(sta.Code) = upper(@state) " +
                                                   "order by  1 ");

                        rs.DoQuery(sql);
                        BPState = rs.Fields.Item("State_Code").Value;
                        BPCounty = rs.Fields.Item("County_AbsId").Value;
                    }
                    catch (Exception)
                    {
                        int errCode;
                        string errMsg;
                        oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception($"{errCode}-{errMsg}");
                    }

                    if (model.Country != "BR")
                    {
                        BP_U_GTTipoPN = "5";
                        BP_U_GToptanteSN = "2";
                        BP_U_SKILL_indIEDest = "9";
                        BP_U_GTorigem = "2";
                    }
                    else if (model.Country == "BR")
                    {
                        if (model.State == "SC")
                        {
                            if (model.CNPJ != null)
                            {
                                BP_U_GTTipoPN = "2";
                                BP_U_GToptanteSN = "2";
                                BP_U_SKILL_indIEDest = "1";
                                BP_U_GTorigem = "1";
                            }
                            else
                            {
                                BP_U_GTTipoPN = "4";
                                BP_U_GToptanteSN = "3";
                                BP_U_SKILL_indIEDest = "9";
                                BP_U_GTorigem = "1";
                            }
                        }
                        else if (model.State == "MG" | model.State == "PR" | model.State == "RS" | model.State == "RJ" | model.State == "SP")
                        {
                            if (model.CNPJ != null)
                            {
                                BP_U_GTTipoPN = "1";
                                BP_U_GToptanteSN = "2";
                                BP_U_SKILL_indIEDest = "1";
                                BP_U_GTorigem = "2";
                            }
                            else
                            {
                                BP_U_GTTipoPN = "3";
                                BP_U_GToptanteSN = "3";
                                BP_U_SKILL_indIEDest = "9";
                                BP_U_GTorigem = "2";
                            }
                        }
                        else
                        {
                            if (model.CNPJ != null)
                            {
                                BP_U_GTTipoPN = "1";
                                BP_U_GToptanteSN = "2";
                                BP_U_SKILL_indIEDest = "1";
                                BP_U_GTorigem = "2";
                            }
                            else
                            {
                                BP_U_GTTipoPN = "3";
                                BP_U_GToptanteSN = "3";
                                BP_U_SKILL_indIEDest = "9";
                                BP_U_GTorigem = "2";
                            }
                        }
                    }
                    else
                    {
                        BP_U_GTTipoPN = "3";
                        BP_U_GToptanteSN = "2";
                        BP_U_SKILL_indIEDest = "9";
                        BP_U_GTorigem = "1";
                    }
                    oBP.CardName = model.CardName;
                    oBP.CardForeignName = model.CardName;
                    oBP.EmailAddress = model.Email;
                    oBP.CardType = BoCardTypes.cCustomer;
                    oBP.SubjectToWithholdingTax = BoYesNoEnum.tNO;
                    oBP.CompanyRegistrationNumber = "1";
                    oBP.UserFields.Fields.Item("U_GTTipoPN").Value = BP_U_GTTipoPN;
                    oBP.UserFields.Fields.Item("U_GToptanteSN").Value = BP_U_GToptanteSN;
                    oBP.UserFields.Fields.Item("U_SKILL_indIEDest").Value = BP_U_SKILL_indIEDest;
                    oBP.UserFields.Fields.Item("U_GTorigem").Value = BP_U_GTorigem;
                    oBP.UserFields.Fields.Item("U_AD_StatusBitrix").Value = "Pendente";
                    oBP.Valid = BoYesNoEnum.tYES;

                    //Bill Address
                    if (LineBilltoDef != 99999)
                    {
                        oBP.Addresses.SetCurrentLine(LineBilltoDef);
                        oBP.Addresses.AddressType = BoAddressType.bo_BillTo;
                        oBP.Addresses.AddressName = "BILL";
                        oBP.Addresses.Street = model.AddressStreet;
                        oBP.Addresses.StreetNo = model.AddressNumber;
                        oBP.Addresses.BuildingFloorRoom = model.AddressComplement;
                        oBP.Addresses.ZipCode = model.ZipCode;
                        oBP.Addresses.Block = model.AddressBlock;
                        oBP.Addresses.City = model.CityName;
                        oBP.Addresses.State = BPState;
                        oBP.Addresses.County = Convert.ToString(BPCounty);
                        oBP.Addresses.Country = model.Country;
                        oBP.Addresses.UserFields.Fields.Item("U_SKILL_indIEDest").Value = BP_U_SKILL_indIEDest;
                    }

                    //Ship Address
                    if (LineShiptoDef != 99999)
                    {
                        oBP.Addresses.SetCurrentLine(LineShiptoDef);
                        oBP.Addresses.AddressType = BoAddressType.bo_ShipTo;
                        oBP.Addresses.AddressName = "SHIP";
                        oBP.Addresses.Street = model.AddressStreet;
                        oBP.Addresses.StreetNo = model.AddressNumber;
                        oBP.Addresses.BuildingFloorRoom = model.AddressComplement;
                        oBP.Addresses.ZipCode = model.ZipCode;
                        oBP.Addresses.Block = model.AddressBlock;
                        oBP.Addresses.City = model.CityName;
                        oBP.Addresses.State = BPState;
                        oBP.Addresses.County = Convert.ToString(BPCounty);
                        oBP.Addresses.Country = model.Country;
                        oBP.Addresses.UserFields.Fields.Item("U_SKILL_indIEDest").Value = BP_U_SKILL_indIEDest;
                    }

                    //Fiscal
                    oBP.FiscalTaxID.SetCurrentLine(0);
                    if (model.CNPJ != null)
                    {
                        oBP.FiscalTaxID.TaxId0 = model.CNPJ;
                        oBP.FiscalTaxID.TaxId4 = "";
                        oBP.FiscalTaxID.TaxId5 = "";
                        if (oBP.FiscalTaxID.TaxId1 == null)
                        {
                            oBP.FiscalTaxID.TaxId1 = "Isento";
                        }
                        if (oBP.FiscalTaxID.CNAECode == 0)
                        {
                            oBP.FiscalTaxID.CNAECode = 255;
                        }
                    }
                    else if (model.CPF != null)
                    {
                        oBP.FiscalTaxID.TaxId0 = "";
                        oBP.FiscalTaxID.TaxId4 = model.CPF;
                        oBP.FiscalTaxID.TaxId5 = "";
                    }
                    else if (model.TaxId != null)
                    {
                        oBP.FiscalTaxID.TaxId0 = "";
                        oBP.FiscalTaxID.TaxId4 = "";
                        oBP.FiscalTaxID.TaxId5 = model.TaxId;
                    }

                    //Payment Method
                    Recordset rs_payment = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    try
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
                    }
                    catch (Exception)
                    {
                        int errCode;
                        string errMsg;
                        oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception($"{errCode}-{errMsg}");
                    }
                    oBP.PeymentMethodCode = "Paypal";

                    //Update
                    if (oBP.Update() != 0)
                    {
                        int errCode;
                        string errMsg;
                        oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception($"{errCode}-{errMsg}");
                    }
                }
                return oBP.CardCode;
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                throw new Exception($"{errCode}-{errMsg}");
            }
        }
        internal static string delete_BP(Company oCompany, B1BusinessPartner model)
        {
            BusinessPartners oBP = (BusinessPartners)oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);

            try
            {
                if (oBP.GetByKey(model.CardCode))
                {
                    oBP.Remove();
                }
                return oBP.CardCode;
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