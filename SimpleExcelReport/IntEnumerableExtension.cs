using System.Collections.Generic;
using System.Linq;

namespace SimpleExcelReport
{
    public static class IntEnumerableExtension
    {
        public static bool Contiguous(this IEnumerable<int> integers)
        {
            List <int> list = integers.ToList();
            list.Sort();

            return list.Contiguous();
        }

        public static bool Contiguous(this List<int> integers)
        {
            int last = integers.First();

            foreach (int i in integers)
            {
                if (i - last > 1)
                {
                    return false;
                }

                last = i;
            }

            return true;
        }
    }
}
