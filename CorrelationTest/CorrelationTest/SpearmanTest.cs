using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CorrelationLib;

namespace CorrelationTest
{
    [TestClass]
    public class SpearmanTest
    {
        [TestMethod]
        [ExpectedException(typeof(System.ArgumentOutOfRangeException))]
        public void Bad1()
        {
            CorrelationCalc correlation = new CorrelationCalc();
            double[] dependents = new double[] { 106, 86, 100, 101, 99, 103, 97, 113, 112 };
            double[] independents = new double[] { 7, 0, 27, 50, 28, 29, 20, 12, 6, 17 };
            correlation.SetDependents(dependents);
            correlation.AddIndependents(independents);
            double coef = Math.Round(correlation.GetSpearmans()[0], 6);
        }

        [TestMethod]
        [ExpectedException(typeof(System.ArgumentOutOfRangeException))]
        public void Bad2()
        {
            CorrelationCalc correlation = new CorrelationCalc();
            double[] dependents = new double[] { 106, 86, 100, 101, 99, 103, 97, 113, 112, 110 };
            double[] independents = new double[] { 7, 0, 27, 50, 28, 29, 20, 12, 6 };
            correlation.SetDependents(dependents);
            correlation.AddIndependents(independents);
            double coef = Math.Round(correlation.GetSpearmans()[0], 6);
        }

        [TestMethod]
        [ExpectedException(typeof(System.Exception))]
        public void Bad3()
        {
            CorrelationCalc correlation = new CorrelationCalc();
            double[] dependents = new double[] { 106, 86, 100, 101, 99, 103, 97, 113, 112, 110 };
            double[] independents = new double[] { 7, 0, 27, 50, 28, 29, 20, 12, 6, 17 };
            correlation.AddIndependents(independents);
            double coef = Math.Round(correlation.GetSpearmans()[0], 6);
        }

        [TestMethod]
        public void Bad4()
        {
            CorrelationCalc correlation = new CorrelationCalc();
            double[] dependents = new double[] { 106, 86, 100, 101, 99, 103, 97, 113, 112, 110 };
            double[] independents = new double[] { 7, 0, 27, 50, 28, 29, 20, 12, 6, 17 };
            correlation.SetDependents(dependents);

            Assert.AreEqual(0, correlation.GetSpearmans().Length);
        }

        [TestMethod]
        public void GetSpearman1()
        {
            CorrelationCalc correlation = new CorrelationCalc();
            double[] dependents = new double[] { 106, 86, 100, 101, 99, 103, 97, 113, 112, 110 };
            double[] independents = new double[] { 7, 0, 27, 50, 28, 29, 20, 12, 6, 17 };
            correlation.SetDependents(dependents);
            correlation.AddIndependents(independents);
            double coef = Math.Round(correlation.GetSpearmans()[0], 6);

            Assert.AreEqual(-0.175758, coef);
        }

        [TestMethod]
        public void GetSpearman2()
        {
            CorrelationCalc correlation = new CorrelationCalc();
            double[] independents = new double[] { 106, 86, 100, 101, 99, 103, 97, 113, 112, 110 };
            double[] dependents = new double[] { 7, 0, 27, 50, 28, 29, 20, 12, 6, 17 };
            correlation.SetDependents(dependents);
            correlation.AddIndependents(independents);
            double coef = Math.Round(correlation.GetSpearmans()[0], 6);

            Assert.AreEqual(-0.175758, coef);
        }
    }
}
