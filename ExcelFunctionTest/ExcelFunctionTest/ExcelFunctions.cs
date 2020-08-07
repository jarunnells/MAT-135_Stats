using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;


//
//   Developer: J.A. Runnells
//     Updated: 2020-07-22 00:45
//   Last Push: 2020-07-22 00:45
//      Branch: master
//     License: MIT
//

// =============================================================================
// AVAILABLE FUNCTIONS (C#):
//   [01] SayHello() <=> [x] TESTING  [x] DEV COMPLETE  **TESTING ONLY**
//   [02] MINUS() <=> [x] TESTING  [ ] DEV COMPLETE
//   [03] PHAT() <=> [x] TESTING  [ ] DEV COMPLETE
//   [04] STDEV_PHAT() <=> [x] TESTING  [ ] DEV COMPLETE
//   [05] MARGINERROR_P() <=> [x] TESTING  [ ] DEV COMPLETE
//   [06] MARGINERROR_M() <=> [x] TESTING  [ ] DEV COMPLETE
//   [07] CONFIDENCE_INTERVAL() <=> [x] TESTING  [ ] DEV COMPLETE
//   [08] TestDefault() <=> [x] TESTING  [ ] DEV COMPLETE
//
//   [**] Internal (helper) Class Optional <=> [x] TESTING  [ ] DEV COMPLETE
// =============================================================================

namespace ExcelFunctionTest
{
    public static class ExcelFunctions
    {
        [ExcelFunction(Description = ".NET Function Test - Hello...")]
        public static string SayHello([ExcelArgument(Name = "name", Description = "A name...")] string name)
        { return $"Hello {name}!!"; }

        [ExcelFunction(Description = ".NET Function Test - AddNums(n1,n2)")]
        public static double MINUS(
            [ExcelArgument(Name = "num1", Description = "Number to subtract from.")] double n1,
            [ExcelArgument(Name = "num2", Description = "Number being subtracted.")] double n2)
        { return n1 - n2; }

        [ExcelFunction(Description = "Calculates p̂ (population proportion) from given X and n values.")]
        public static double PHAT(
            [ExcelArgument(Name = "x", Description = "individuals in the sample with a specified characteristic")] double x,
            [ExcelArgument(Name = "n", Description = "sample size")] double n)
        { return x / n; }

        // TODO: complete descriptions
        [ExcelFunction(Description ="Calculates standard deviation of phat")]
        public static double STDEV_PHAT(
            [ExcelArgument(Name = "phat", Description = "phat value description")] double phat,
            [ExcelArgument(Name = "n", Description = "sample size description")] double n)
        { return Math.Sqrt(phat * (1 - phat) / n); }

        // TODO: complete descriptions
        [ExcelFunction(Description = "Calculates the margin of error [E or ME] for a proportion.")] 
        public static double MARGINERROR_P(
            [ExcelArgument(Name = "z_alpha2", Description ="")] double z_alpha2,
            [ExcelArgument(Name = "phat", Description ="")] double phat,
            [ExcelArgument(Name = "n", Description ="")] double n)
        { return z_alpha2 * Math.Sqrt(phat * (1 - phat) / n); }

        // TODO: complete descriptions
        [ExcelFunction(Description = "Calculates the margin of error [E or ME] for a proportion.")]
        public static double MARGINERROR_M(
            [ExcelArgument(Name = "t_alpha2", Description = "")] double t_alpha2,
            [ExcelArgument(Name = "s", Description = "")] double s,
            [ExcelArgument(Name = "n", Description = "")] double n)
        { return t_alpha2 * (s / Math.Sqrt(n)); }

        // TODO: complete switch/case implementation
        [ExcelFunction(Description = "Calculates the margin of error [E or ME] for a proportion.")]
        public static double CONFIDENCE_INTERVAL(
            [ExcelArgument(Name = "z_alpha2", Description = "")] double t_alpha2,
            [ExcelArgument(Name = "phat", Description = "")] double s,
            [ExcelArgument(Name = "n", Description = "")] double n,
            [ExcelArgument(Name = "lower", Description = "")] bool lower = false)
        { return 0; }

        // TODO: complete descriptons
        [ExcelFunction(Description = "Function to test Optional helper class implementation.")]
        public static string TestDefault(
            [ExcelArgument(Name = "nameArg", Description = "")] object nameArg)
        { return $"Hello {Optional.Check(nameArg, " Unknown Person!?")}"; }
    }

    // Helper Class - Optionals
    internal static class Optional
    {
        internal static bool Check(object arg, bool defaultVal)
        { return defaultVal; }

        internal static int Check(object arg, int defaultVal)
        {
            if (arg is int)
                return (int)arg;
            else if (arg is ExcelMissing)
                return defaultVal;
            else
                throw new ArgumentException();
        }

        internal static double Check(object arg, double defaultVal)
        {
            if (arg is double)
                return (double)arg;
            else if (arg is ExcelMissing)
                return defaultVal;
            else
                throw new ArgumentException();
        }

        // TODO: add more type checks 
        internal static string Check(object arg, string defaultVal)
        {
            if (arg is string)
                return (string)arg;
            else if (arg is ExcelMissing)
                return defaultVal;
            else
                return arg.ToString();
        }

        internal static DateTime Check(object arg, DateTime defaultVal)
        {
            if (arg is double)
                return DateTime.FromOADate((double)arg);
            else if (arg is string)
                return DateTime.Parse((string)arg);
            else if (arg is ExcelMissing)
                return defaultVal;
            else
                throw new ArgumentException();
        }
    }
}
