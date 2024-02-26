using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Links
{
    class Media
    {
        public int poz = 0;
        public int neg = 0;
        public int neu = 0;
        public string name = "";
        public string link = "";

        public static bool operator >(Media first, Media second)
        {
            if (first.neg < second.neg)
                return false;
            if (first.neg > second.neg)
                return true;
            if ((first.neg == second.neg) && (first.poz > second.poz))
                return true;
            if ((first.neg == second.neg) && (first.poz < second.poz))
                return false;
            if ((first.neg == second.neg) && (first.poz == second.poz) && (first.neu > second.neu))
                return true;
            if ((first.neg == second.neg) && (first.poz == second.poz) && (first.neu <= second.neu))
                return false;
            return false;
        }

        public static bool operator <(Media first, Media second)
        {
            return second > first;
        }

    }
}
