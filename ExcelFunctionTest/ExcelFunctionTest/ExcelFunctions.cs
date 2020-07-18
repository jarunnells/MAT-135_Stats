using System;
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
        public static double AddNums(
            [ExcelArgument(Name = "num1", Description = "First number to add")] double n1,
            [ExcelArgument(Name = "num2", Description = "Second number to add")] double n2)
        {
            return n1 + n2;
        }

        [ExcelFunction(Description = "Calculates p̂ (population proportion) from given X and n values.")]
        public static double PHAT(
            [ExcelArgument(Name = "x", Description = "individuals in the sample with a specified characteristic")] double x,
            [ExcelArgument(Name = "n", Description = "sample size")] double n)
        {
            return x / n;
        }
    }
}
