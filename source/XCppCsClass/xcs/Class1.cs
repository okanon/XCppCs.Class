using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.InteropServices;

namespace xcs
{
    /// <summary>
    /// Interface
    /// </summary>
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class XCsClass
    {
        public Int32 Sum(Int32 num1, Int32 num2)
        {
            return num1 + num2;
        }

        public Int32 Multi(Int32 num1, Int32 num2)
        {
            return num1 * num2;
        }

        public string Name(string name, int age)
        {
            return String.Format("My name is {0}. Age {1}", name, age);
        }

        public Int32 Once(Int32 num)
        {
            return num;
        }

        public void Hello()
        {
            Console.WriteLine("Hello use C# COM on C++ World!!");
        }
    }
}
