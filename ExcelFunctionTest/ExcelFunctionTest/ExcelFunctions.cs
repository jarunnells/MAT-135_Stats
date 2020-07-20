﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelFunctionTest
{
    public static class ExcelFunctions
    {
        [ExcelFunction(Description = ".NET Function Test - Hello...")]
        public static string SayHello([ExcelArgument(Name = "name", Description = "A name...")] string name)
        {
            return $"Hello {name}";
        }

        [ExcelFunction(Description = ".NET Function Test - AddNums(n1,n2)")]
        public static double MINUS(
            [ExcelArgument(Name = "num1", Description = "Number to subtract from.")] double n1,
            [ExcelArgument(Name = "num2", Description = "Number being subtracted.")] double n2)
        {
            return n1 - n2;
        }

        [ExcelFunction(Description = "Calculates p̂ (population proportion) from given X and n values.")]
        public static double PHAT(
            [ExcelArgument(Name = "x", Description = "individuals in the sample with a specified characteristic")] double x,
            [ExcelArgument(Name = "n", Description = "sample size")] double n)
        {
            return x / n;
        }

        [ExcelFunction(Description ="Calculates standard deviation of phat")]
        public static double STDEV_PHAT(
            [ExcelArgument(Name = "phat", Description = "phat value description")] double phat,
            [ExcelArgument(Name = "n", Description = "sample size description")] double n)
        {
            return Math.Sqrt(phat * (1 - phat) / n);
        }
    }
}