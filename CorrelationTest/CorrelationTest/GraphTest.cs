using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CorrelationLib;

namespace CorrelationTest
{
    [TestClass]
    public class GraphTest
    {
        [TestMethod]
        public void CreateGraphs()
        {
            string dir = String.Empty;
            CorrelationCalc correlation = new CorrelationCalc();
            double[] dependents = new double[] { 43, 21, 25, 42, 57, 59 };
            double[] independents = new double[] { 99, 65, 79, 75, 87, 81 };
            double[] independents2 = new double[] { 65, 99, 81, 87, 75, 79 };
            correlation.SetDependents(dependents, "IQ");
            correlation.AddIndependents(independents, "reading time");
            correlation.AddIndependents(independents2, "sleeping time");
            correlation.MakeGraphs(dir, true, true);
        }
    }
}
