using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.Mvc;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using InsightsaddinWebApp;
using InsightsaddinWebApp.Controllers;

namespace InsightsaddinWebApp.Tests.Controllers
{
    [TestClass]
    public class PartnerAccountCorntrollerTest
    {
        [TestMethod]
        public void Index()
        {
            // Arrange
            PartnerAccountController controller = new PartnerAccountController();

            // Act
            ViewResult result = controller.Index() as ViewResult;

            // Assert
            Assert.IsNotNull(result);
        }
    }
}
