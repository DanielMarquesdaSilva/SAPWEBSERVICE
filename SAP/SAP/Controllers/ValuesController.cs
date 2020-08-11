using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Claims;
using System.Web.Http;
using SAP.Models;
using SAPbobsCOM;
using AuthorizeAttribute = SAP.Models.AuthorizeAttribute;

namespace SAP.Controllers
{
    public class ValuesController : ApiController
    {
        //http://localhost/SAP/

        //*************CONECT WS TOKEN AND SECURITY*************
        [AllowAnonymous]
        [HttpGet]
        [Route("api/data/forall")]
        public IHttpActionResult Get()
        {
            return Ok("Now server time is: " + DateTime.Now.ToString());
        }

        [Authorize]
        [HttpGet]
        [Route("api/data/authenticate")]
        public IHttpActionResult GetforAuthenticate()
        {
            var identity = (ClaimsIdentity)User.Identity;
            return Ok("Hello " + identity.Name);
        }

        [Authorize(Roles = "admin")]
        [HttpGet]
        [Route("api/data/authorize")]
        public IHttpActionResult GetForAdmin()
        {
            var identity = (ClaimsIdentity)User.Identity;
            var roles = identity.Claims
                        .Where(c => c.Type == ClaimTypes.Role)
                        .Select(c => c.Value);
            return Ok("Hello " + identity.Name + " Role: " + string.Join(",", roles.ToList()));
        }

        //SAP

        //*************CONNECT SAP DLL*************
        [Authorize]
        [HttpPost]
        [Route("api/data/connection")]
        public IHttpActionResult PostConnection([FromBody]B1Connection model)
        {
            string message = "";
            try
            {
                oCompany = new Company();
                B1Connection.create_CN(oCompany, model);
                return Ok(Convert.ToString(Convert.ToString(oCompany.GetCompanyTime()) + " - " + Convert.ToString(oCompany.Connected)));
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                message = message + " | Error SAP: " + errCode.ToString() + "-" + errMsg.ToString();
                var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new StringContent(string.Format("{0}", message)),
                    ReasonPhrase = "Connection error"
                };
                throw new HttpResponseException(resp);
            }
        }

        [Authorize]
        [HttpGet]
        [Route("api/data/connection")]
        public IHttpActionResult GetConnection()
        {
            string message = "";
            try
            {
                return Ok(B1Connection.read_CN(oCompany));
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                message = message + " | Error SAP: " + errCode.ToString() + "-" + errMsg.ToString();
                var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new StringContent(string.Format("{0}", message)),
                    ReasonPhrase = "Connection error"
                };
                throw new HttpResponseException(resp);
            }
        }

        [Authorize]
        [HttpPut]
        [Route("api/data/connection")]
        public IHttpActionResult PutConnection()
        {
            return Ok("Put");
        }

        [Authorize]
        [HttpDelete]
        [Route("api/data/connection")]
        public IHttpActionResult DeleteConnection()
        {
            return Ok("Delete");
        }

        //*************SAP BUSINESS PARTNER*************

        [Authorize]
        [HttpPost]
        [Route("api/data/businesspartner")]
        public IHttpActionResult PostBusinessPartner([FromBody]B1BusinessPartner model)
        {
            string message = "";
            try
            {
                if (model.CardName == null | model.Email == null | (model.CNPJ == null & model.CPF == null & model.TaxId == null) | model.AddressStreet == null | model.ZipCode == null | model.CityName == null | model.State == null | model.Country == null)
                {
                    message = message + "Missing information: ";
                    if (model.CardName == null)
                    {
                        message = message + "/CardName";
                    }
                    if (model.Email == null)
                    {
                        message = message + "/Email";
                    }
                    if (model.CNPJ == null)
                    {
                        message = message + "/CNPJ";
                    }
                    if (model.CPF == null)
                    {
                        message = message + "/CPF";
                    }
                    if (model.TaxId == null)
                    {
                        message = message + "/TaxId";
                    }
                    if (model.AddressStreet == null)
                    {
                        message = message + "/AddressStreet";
                    }
                    if (model.ZipCode == null)
                    {
                        message = message + "/ZipCode";
                    }
                    if (model.CityName == null)
                    {
                        message = message + "/CityName";
                    }
                    if (model.State == null)
                    {
                        message = message + "/State";
                    }
                    if (model.Country == null)
                    {
                        message = message + "Country";
                    }
                    var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                    {
                        Content = new StringContent(string.Format("{0}", message)),
                        ReasonPhrase = "CardCode not created"
                    };
                    throw new HttpResponseException(resp);
                }
                else
                {
                    return Ok(B1BusinessPartner.create_BP(oCompany, model));
                }
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                message = message + " | Error SAP: " + errCode.ToString() + "-" + errMsg.ToString();
                var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new StringContent(string.Format("{0}", message)),
                    ReasonPhrase = "CardCode not created"
                };
                throw new HttpResponseException(resp);
            }
        }

        [Authorize]
        [HttpGet]
        [Route("api/data/businesspartner")]
        public IHttpActionResult GetBusinessPartner([FromBody]B1BusinessPartner model)
        {
            string message = "";
            string CardCode = "";
            try
            {
                CardCode = B1BusinessPartner.read_BP(oCompany, model);
                if (CardCode == null | CardCode == "")
                {
                    var resp = new HttpResponseMessage(HttpStatusCode.NotFound)
                    {
                        Content = new StringContent(string.Format("{0}", message)),
                        ReasonPhrase = "CardCode not found"
                    };
                    throw new HttpResponseException(resp);
                }
                else
                {
                    return Ok(CardCode.ToString());
                }
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                message = message + " | Error SAP: " + errCode.ToString() + "-" + errMsg.ToString();
                var resp = new HttpResponseMessage(HttpStatusCode.NotFound)
                {
                    Content = new StringContent(string.Format("{0}", message)),
                    ReasonPhrase = "CardCode not found"
                };
                throw new HttpResponseException(resp);
            }
        }

        [Authorize]
        [HttpGet] //Get with ID for Hotmart development from IT
        [Route("api/data/businesspartnerid")]
        public IHttpActionResult GetBusinessPartner(string CardCode, string CNPJ, string CPF, string TaxId)
        {
            string message = "";
            try
            {
                if (CardCode == null & CNPJ == null & CPF == null & TaxId == null)
                {
                    var resp = new HttpResponseMessage(HttpStatusCode.NotFound)
                    {
                        Content = new StringContent(string.Format("{0}", message)),
                        ReasonPhrase = "CardCode not found"
                    };
                    throw new HttpResponseException(resp);
                }
                else
                {
                    var model = new B1BusinessPartner();
                    if (CNPJ != null)
                    {
                        model.CNPJ = CNPJ;
                    }
                    else if (CPF != null)
                    {
                        model.CPF = CPF;
                    }
                    else if (CardCode != null)
                    {
                        model.CardCode = CardCode;
                    }
                    else
                    {
                        model.TaxId = TaxId;
                    }
                    CardCode = B1BusinessPartner.read_BP(oCompany, model);
                    if (CardCode == null | CardCode == "")
                    {
                        var resp = new HttpResponseMessage(HttpStatusCode.NotFound)
                        {
                            Content = new StringContent(string.Format("{0}", message)),
                            ReasonPhrase = "CardCode not found"
                        };
                        throw new HttpResponseException(resp);
                    }
                }
                return Ok(CardCode);
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                message = message + " | Error SAP: " + errCode.ToString() + "-" + errMsg.ToString();
                var resp = new HttpResponseMessage(HttpStatusCode.NotFound)
                {
                    Content = new StringContent(string.Format("{0}", message)),
                    ReasonPhrase = "CardCode not found"
                };
                throw new HttpResponseException(resp);
            }
        }

        [Authorize]
        [HttpPut]
        [Route("api/data/businesspartner")]
        public IHttpActionResult PutBusinessPartner([FromBody]B1BusinessPartner model)
        {
            string message = "";
            try
            {
                if (model.CardCode == null)
                {
                    message = message + "Missing information: ";
                    if (model.CardCode == null)
                    {
                        message = message + "/CardCode";
                    }
                    var resp = new HttpResponseMessage(HttpStatusCode.NotFound)
                    {
                        Content = new StringContent(string.Format("{0}", message)),
                        ReasonPhrase = "CardCode was not found"
                    };
                    throw new HttpResponseException(resp);
                }
                else
                {
                    return Ok(B1BusinessPartner.update_BP(oCompany, model));
                }
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                message = message + " | Error SAP: " + errCode.ToString() + "-" + errMsg.ToString();
                var resp = new HttpResponseMessage(HttpStatusCode.NotFound)
                {
                    Content = new StringContent(string.Format("{0}", message)),
                    ReasonPhrase = "CardCode was not found"
                };
                throw new HttpResponseException(resp);
            }
        }

        [Authorize]
        [HttpDelete]
        [Route("api/data/businesspartner")]
        public IHttpActionResult DeleteBusinessPartner([FromBody]B1BusinessPartner model)
        {
            string message = "";
            try
            {
                if (model.CardCode == "")
                {
                    message = message + "Missing information: ";
                    var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                    {
                        Content = new StringContent(string.Format("{0}",message)),
                        ReasonPhrase = "CardCode not created"
                    };
                    throw new HttpResponseException(resp);
                }
                else
                {
                    return Ok(B1BusinessPartner.delete_BP(oCompany, model));
                }
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                message = message + " | Error SAP: " + errCode.ToString() + "-" + errMsg.ToString();
                var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new StringContent(string.Format("{0}", message)),
                    ReasonPhrase = "CardCode was not found"
                };
                throw new HttpResponseException(resp);
            }
        }

        //*************SAP BLANKET AGREEMENT*************

        [Authorize]
        [HttpPost]
        [Route("api/data/blanketagreement")]
        public IHttpActionResult PostBlanketAgreement([FromBody]B1SalesBlanketAgreement model)
        {
            string getKey = "";
            string message = "";
            try
            {
                if ((model.Origin != "Paypal" && model.Origin != "Hotmart") | (model.Currency != "BRL" && model.ValueRate == 0.00) | ((model.GrossAmount - model.FeeAmount - model.NetAmount) != 0.00) | model.CardCode == null | model.AssignSubscriptionId == null | model.AssignSubscriptionTransactionId == null | model.InitialDate == null | model.Currency == null)
                {
                    message = "Missing information: ";
                    if (model.Origin != "Paypal" && model.Origin != "Hotmart")
                    {
                        message = message + "/Origin have to be Paypal or Hotmart";
                    }
                    if (model.Currency != "BRL" && model.ValueRate == 0.00)
                    {
                        message = message + "/If currency is different of BRL you have to inform ValueRate";
                    }
                    if ((model.GrossAmount - model.FeeAmount - model.NetAmount) != 0.00)
                    {
                        message = message + "/Sum GrossAmount less FeeAmount ples NetAmount have to be zero";
                    }
                    if (model.CardCode == null)
                    {
                        message = message + "/CardCode";
                    }
                    if (model.AssignSubscriptionId == null)
                    {
                        message = message + "/AssignSubscriptionId";
                    }
                    if (model.AssignSubscriptionTransactionId == null)
                    {
                        message = message + "/AssignSubscriptionTransactionId";
                    }
                    if (model.InitialDate == null)
                    {
                        message = message + "/InitialDate";
                    }
                    if (model.Currency == null)
                    {
                        message = message + "/Currency";
                    }
                    var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                    {
                        Content = new StringContent(string.Format("{0}", message)),
                        ReasonPhrase = "Blancket agreement not created"
                    };
                    throw new HttpResponseException(resp);
                }
                else
                {
                    if (model.Origin == "Paypal")
                    {
                        getKey = B1SalesBlanketAgreement.create_BA(oCompany, model) + "|";
                        getKey = getKey + B1JournalEntry.create_JE(oCompany, model) + "|";
                        getKey = getKey + B1SalesInvoice.create_IV(oCompany, model);
                    }
                    else if(model.Origin == "Hotmart")
                    {
                       getKey = B1SalesBlanketAgreement.create_BA(oCompany, model) + "|";
                       getKey = getKey + B1SalesInvoice.create_IV(oCompany, model);
                    }
                    return Ok(getKey);
                }
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                message = message + " | Error SAP: " + errCode.ToString() + "-" + errMsg.ToString();
                var resp = new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new StringContent(string.Format("{0}", message)),
                    ReasonPhrase = "Blancket agreement not created"
                };
                throw new HttpResponseException(resp);
            }
        }
        static Company oCompany;
    }
}
