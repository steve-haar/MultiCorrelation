using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CorrelationLib;

namespace CorrelationTest
{
    [TestClass]
    public class PearsonTest
    {
        [TestMethod]
        [ExpectedException(typeof(System.ArgumentOutOfRangeException))]
        public void Bad1()
        {
            CorrelationCalc correlation = new CorrelationCalc();
            double[] dependents = new double[] { 43, 21, 25, 42, 57 };
            double[] independents = new double[] { 99, 65, 79, 75, 87, 81 };
            correlation.SetDependents(dependents);
            correlation.AddIndependents(independents);
            double coef = Math.Round(correlation.GetPearsons()[0], 6);
        }

        [TestMethod]
        [ExpectedException(typeof(System.ArgumentOutOfRangeException))]
        public void Bad2()
        {
            CorrelationCalc correlation = new CorrelationCalc();
            double[] dependents = new double[] { 43, 21, 25, 42, 57, 59 };
            double[] independents = new double[] { 99, 65, 79, 75, 87 };
            correlation.SetDependents(dependents);
            correlation.AddIndependents(independents);
            double coef = Math.Round(correlation.GetPearsons()[0], 6);
        }

        [TestMethod]
        [ExpectedException(typeof(System.Exception))]
        public void Bad3()
        {
            CorrelationCalc correlation = new CorrelationCalc();
            double[] dependents = new double[] { 43, 21, 25, 42, 57, 59 };
            double[] independents = new double[] { 99, 65, 79, 75, 87 };
            correlation.AddIndependents(independents);
            double coef = Math.Round(correlation.GetPearsons()[0], 6);
        }

        [TestMethod]
        public void Bad4()
        {
            CorrelationCalc correlation = new CorrelationCalc();
            double[] dependents = new double[] { 43, 21, 25, 42, 57, 59 };
            double[] independents = new double[] { 99, 65, 79, 75, 87 };
            correlation.SetDependents(dependents);

            Assert.AreEqual(0, correlation.GetPearsons().Length);
        }

        [TestMethod]
        public void GetPearson1()
        {
            CorrelationCalc correlation = new CorrelationCalc();
            double[] dependents = new double[] { 43, 21, 25, 42, 57, 59 };
            double[] independents = new double[] { 99, 65, 79, 75, 87, 81 };
            correlation.SetDependents(dependents);
            correlation.AddIndependents(independents);
            double coef = Math.Round(correlation.GetPearsons()[0], 6);

            Assert.AreEqual(0.529809, coef);
        }

        [TestMethod]
        public void GetPearson2()
        {
            CorrelationCalc correlation = new CorrelationCalc();
            double[] independents = new double[] { 43, 21, 25, 42, 57, 59 };
            double[] dependents = new double[] { 99, 65, 79, 75, 87, 81 };
            correlation.SetDependents(dependents);
            correlation.AddIndependents(independents);
            double coef = Math.Round(correlation.GetPearsons()[0], 6);

            Assert.AreEqual(0.529809, coef);
        }
    }
}
