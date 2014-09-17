using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.Framework.Client;

namespace ReleaseNotes
{
    [TestClass]
    public class UnitTests
    {
        [TestMethod]
        public void testConnection()
        {
            TFSAccessor t = TFSAccessor.TFSAccessorFactory();
            Assert.IsNotNull(t);
        }

        [TestMethod]
        public void testQuery()
        {
            TFSAccessor t = TFSAccessor.TFSAccessorFactory();
            WorkItemCollection wic = t.getReleaseNotesFromQuery("Dealspan", "14.4");
            Assert.IsNotNull(wic);
        }

        [TestMethod]
        public void testSpacing()
        {
            Assert.AreEqual(Utilities.spaceCapitalizedNames("HelloWorldOfAwesomeStuff"), "Hello World Of Awesome Stuff");
        }
    }
}

