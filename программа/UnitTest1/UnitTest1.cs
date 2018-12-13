using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WindowsFormsApplication1;

namespace UnitTest1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            string str = "";
            bool ogedaem = false;
            Form1 t = new Form1();
            bool rezult = t.Date(str);
            Assert.AreEqual(ogedaem,rezult);

        }

        [TestMethod]
        public void TestMethod2()
        {
            string str = "12.12.2017";
            bool ogedaem = true;
            Form1 t = new Form1();
            bool rezult = t.Date(str);
            Assert.AreEqual(ogedaem, rezult);

        }


        [TestMethod]
        public void TestMethod3()
        {
            string str = "12.13.2017";
            bool ogedaem = false;
            Form1 t = new Form1();
            bool rezult = t.Date(str);
            Assert.AreEqual(ogedaem, rezult);

        }
    }
}
