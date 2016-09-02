using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using InsightsaddinWebApp.Models;
using InsightsaddinWebApp.Controllers;
using System.Threading.Tasks;
using System.Net;
using Microsoft.ServiceBus.Messaging;
using System.Text;
using System.Threading;

namespace InsightsaddinWebApp.Controllers
{
    public class PartnerAccountController : Controller
    {
        static string eventHubName = "insightsaddin-eh"; //"{Event Hub name}"
        static string connectionString = "";// "{send connection string}"

        public void SendEventHubMessage()
        {
            var eventHubClient = EventHubClient.CreateFromConnectionString(connectionString, eventHubName);
            while (true)
            {
                try
                {
                    var message = Guid.NewGuid().ToString();
                    Console.WriteLine("{0} > Sending message: {1}", DateTime.Now, message);
                    eventHubClient.Send(new EventData(Encoding.UTF8.GetBytes(message)));
                }
                catch (Exception exception)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("{0} > Exception: {1}", DateTime.Now, exception.Message);
                    Console.ResetColor();
                }

                Thread.Sleep(200);
            }
        }

        public void GetEventHubMessage()
        {

        }

        public ActionResult Index()
        {
            var partnerAccounts = DocumentDBRepository.GetIncompletePartnerAccounts();
            return this.View(partnerAccounts);
        }

        public ActionResult Create()
        {
            return this.View();
        }

        [HttpPost] 
        [ValidateAntiForgeryToken]
        // http://www.asp.net/web-api/overview/security/preventing-cross-site-request-forgery-csrf-attacks
        // ^ above ^ help protect this application against cross-site request forgery attacks
        public async Task<ActionResult> Create([Bind(Include = "Pbe, Website, Crm, Stage, EngagementType, Date, Reason, Location, Meeting, Industry, CloudStatus, CloudProvider, Consumption, WorkLoads")] PartnerAccount partnerAccount)
        {
            if (ModelState.IsValid)
            {
                await DocumentDBRepository.CreatePartnerAccountAsync(partnerAccount);
                return this.RedirectToAction("Index");
            }

            Console.WriteLine("Creation in progress...");



            return this.View(partnerAccount);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Edit([Bind(Include = "Pbe, Website, Crm, Stage, EngagementType, Date, Reason, Location, Meeting, Industry, CloudStatus, CloudProvider, Consumption, WorkLoads")] PartnerAccount partnerAccount)
        {
            if (ModelState.IsValid)
            {
                await DocumentDBRepository.UpdatePartnerAccountAsync(partnerAccount);
                return this.RedirectToAction("Index");
            }

            return this.View(partnerAccount);
        }

        public ActionResult Edit(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            PartnerAccount partnerAccount = (PartnerAccount) DocumentDBRepository.GetPartnerAccount(id);
            if (partnerAccount == null)
            {
                return this.HttpNotFound();
            }

            return this.View(partnerAccount);
        }
    }
}