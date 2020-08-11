using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SAPbobsCOM;

namespace SAP.Models
{
    public class B1SalesBlanketAgreement
    {
        public string Origin { get; set; }
        public string CardCode { get; set; }
        public string AssignSubscriptionId { get; set; }
        public string AssignSubscriptionTransactionId { get; set; }
        public DateTime InitialDate { get; set; }
        public string Currency { get; set; }
        public double GrossAmount { get; set; }
        public double FeeAmount { get; set; }
        public double NetAmount { get; set; }
        public double ValueRate { get; set; }
        public string SellerEmail { get; set; }
        public string CouponCode { get; set; }
        public string ItemCode { get; set; }

        internal static string create_BA(Company oCompany, B1SalesBlanketAgreement model)
        {
            BlanketAgreementsService oBAService = (BlanketAgreementsService)oCompany.GetCompanyService().GetBusinessService(ServiceTypes.BlanketAgreementsService);
            BlanketAgreement oBA = (BlanketAgreement)oBAService.GetDataInterface(BlanketAgreementsServiceDataInterfaces.basBlanketAgreement);
            BlanketAgreements_ItemsLine itemLine;

            try
            {
                oBA.BPCode = model.CardCode;
                oBA.StartDate = model.InitialDate;
                oBA.EndDate = model.InitialDate.AddYears(1);
                oBA.Description = model.Origin;
                oBA.Status = BlanketAgreementStatusEnum.asApproved;
                oBA.PriceMode = PriceModeEnum.pmNet;
                oBA.BPCurrency = model.Currency;
                oBA.AgreementMethod = BlanketAgreementMethodEnum.amItem;
                oBA.Remarks = "Agreement from " + model.Origin + " subscription: " + model.AssignSubscriptionId + "Coupon code = " + model.CouponCode + " Seller mail: " + model.SellerEmail;
                oBA.UserFields.Item("U_SubscriptionId").Value = model.AssignSubscriptionId;
                itemLine = oBA.BlanketAgreements_ItemsLines.Add();
                itemLine.ItemNo = model.ItemCode;
                itemLine.PlannedQuantity = 1;
                if (model.Currency == "USD" | model.Currency == "EUR")
                {
                    itemLine.PriceCurrency = model.Currency;
                }
                itemLine.UnitPrice = model.GrossAmount;
                oBAService.AddBlanketAgreement(oBA);
                return read_BA(oCompany, model).ToString();
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                throw new Exception($"{errCode}-{errMsg}");
            }
        }
        internal static int read_BA(Company oCompany, B1SalesBlanketAgreement model)
        {
            Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            int BlancketNumber;
            try
            {
                string sql = string.Format("select    top 1 con.AbsID as AbsId                             " +
                                           "from      [dbo].[OOAT] con                                     " +
                                           "          inner join[dbo].[OAT1] ite on(con.AbsID = ite.AgrNo) " +
                                           "where     con.BpCode = '" + model.CardCode + "'                " +
                                           "and       ite.ItemCode =  '" + model.ItemCode + "'             " +
                                           "and       con.StartDate <= getdate() + 30                      " +
                                           "and       con.EndDate >= getdate() - 30                        " +
                                           "and       con.Status <> 'T'                                    " +
                                           "order by  1 desc                                               ");

                rs.DoQuery(sql);
                BlancketNumber = rs.Fields.Item("AbsId").Value;
                return BlancketNumber;
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