using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CorrelationLib;

namespace CorrelationTest
{
    [TestClass]
    public class CorrelationTest
    {
        [TestMethod]
        public void SetDependent()
        {
            CorrelationCalc correlation = new CorrelationCalc();
            double[] dependents = new double[] { };
            correlation.SetDependents(dependents);
        }

        [TestMethod]
        public void AddIndependent()
        {
            CorrelationCalc correlation = new CorrelationCalc();
            double[] independents = new double[] { };
            correlation.AddIndependents(independents);
        }
    }
}
