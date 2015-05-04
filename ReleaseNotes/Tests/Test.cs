using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.Framework.Client;

namespace ReleaseNotesLibrary
{
    [TestClass]
    public class UnitTests
    {
        [TestMethod]
        public void testConnection()
        {
            TFSAccessor t = TFSAccessor.TFSAccessorFactory("", "", "", "");
            Assert.IsNotNull(t);
        }

        [TestMethod]
        public void testQuery()
        {
            TFSAccessor t = TFSAccessor.TFSAccessorFactory("", "", "", "");
            WorkItemCollection wic = t.GetReleaseNotesFromQuery();
            Assert.IsNotNull(wic);
        }

        [TestMethod]
        public void testSpacing()
        {
            Assert.AreEqual(Utilities.SpaceCapitalizedNames("HelloWorldOfAwesomeStuff"), "Hello World Of Awesome Stuff");
        }
    }
}

