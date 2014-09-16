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
        public void TestConnection()
        {
            TFSAccessor t = TFSAccessor.TFSAccessorFactory();
            Assert.IsNotNull(t);
        }

        [TestMethod]
        public void TestQuery()
        {
            TFSAccessor t = TFSAccessor.TFSAccessorFactory();
            WorkItemCollection wic = t.getReleaseNotesFromQuery("Dealspan", "14.4");
            Assert.IsNotNull(wic);
        }
    }
}

